import os
import re
import glob
import threading
import queue
import PyPDF2
import pdfplumber
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import List, Dict, Tuple
from datetime import datetime

class PDFSplitter:
    """PDF分割器核心类"""
    def __init__(self, log_callback=None):
        self.log_callback = log_callback
    
    def log(self, message):
        """记录日志"""
        if self.log_callback:
            self.log_callback(message)
    
    def extract_toc_text(self, pdf_path: str) -> str:
        """提取PDF第二页（目录页）的文本内容"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) >= 2:
                    toc_page = pdf.pages[1]
                    text = toc_page.extract_text()
                    return text
                else:
                    self.log(f"警告: {Path(pdf_path).name} 页数不足2页")
                    return ""
        except Exception as e:
            self.log(f"提取目录页失败 {Path(pdf_path).name}: {str(e)}")
            return ""
    
    def parse_toc(self, toc_text: str) -> List[Dict[str, int]]:
        """解析目录文本，提取章节和页码"""
        sections = []
        
        # 常见的目录格式模式
        patterns = [
            # 模式1: "第一章 标题 ...... 1"
            r'(?:第[一二三四五六七八九十\d]+章|[一二三四五六七八九十\d]+\.\d+|[一二三四五六七八九十\d]+\.)\s*([^……\n]+?)[……\s\.]*(\d+)\s*$',
            
            # 模式2: "第一章 标题 1"
            r'(?:第[一二三四五六七八九十\d]+章|[一二三四五六七八九十\d]+\.\d+)\s+(.+?)\s+(\d+)\s*$',
            
            # 模式3: 英文模式
            r'(?:Chapter|CHAPTER|Part|PART)\s*[\dIVX]+[\.\s]+(.+?)[\.\s]*(\d+)\s*$',
            
            # 模式4: 简单的数字页码模式
            r'(.+?)\s+(\d{1,3})\s*$',
            
            # 模式5: 包含点号的模式
            r'(\d+\.\s*.+?)\s+(\d{1,3})\s*$'
        ]
        
        lines = toc_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 跳过明显的非目录行
            skip_keywords = ['目录', 'CONTENTS', '目次', '===', '---', '...']
            if any(keyword in line for keyword in skip_keywords):
                continue
                
            # 尝试多种模式匹配
            for pattern in patterns:
                match = re.search(pattern, line)
                if match:
                    section_name = match.group(1).strip()
                    try:
                        page_num = int(match.group(2))
                        if 0 < page_num < 1000:
                            sections.append({
                                'name': section_name,
                                'page': page_num
                            })
                            break
                    except ValueError:
                        continue
        
        # 备选方法
        if len(sections) < 3:
            sections = self._fallback_parse_toc(lines)
            
        return sections
    
    def _fallback_parse_toc(self, lines: List[str]) -> List[Dict[str, int]]:
        """备选解析方法"""
        sections = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            match = re.search(r'(\d+)\s*$', line)
            if match:
                page_num = int(match.group(1))
                section_name = line[:match.start()].strip()
                section_name = re.sub(r'[\.\s]+$', '', section_name)
                
                if section_name and 0 < page_num < 1000:
                    sections.append({
                        'name': section_name,
                        'page': page_num
                    })
        
        return sections
    
    def validate_and_adjust_sections(self, sections: List[Dict[str, int]], 
                                   total_pages: int) -> List[Dict[str, int]]:
        """验证并调整章节页码"""
        if not sections:
            return self._auto_split_sections(total_pages)
        
        sections.sort(key=lambda x: x['page'])
        
        # 如果章节数少于5个，自动补全
        if len(sections) < 5:
            self.log("  章节数不足5个，将自动分割")
            return self._auto_split_sections(total_pages)
        
        # 如果章节数大于5个，选择前5个
        if len(sections) > 5:
            selected = [sections[0]]
            for i in range(1, min(6, len(sections))):
                selected.append(sections[i])
            sections = selected
        
        # 确保最后一个章节包含最后一页
        if sections[-1]['page'] < total_pages:
            sections.append({'name': '最后部分', 'page': total_pages})
        
        return sections[:5]
    
    def _auto_split_sections(self, total_pages: int) -> List[Dict[str, int]]:
        """自动将PDF分成5个大致相等的部分"""
        part_size = total_pages // 5
        sections = []
        
        for i in range(5):
            page_num = i * part_size + 1
            if i == 4:
                page_num = total_pages
            sections.append({
                'name': f'第{i+1}部分',
                'page': page_num
            })
        
        return sections
    
    def split_pdf_by_sections(self, pdf_path: str, sections: List[Dict[str, int]], 
                              output_folder: str) -> bool:
        """根据章节信息分割PDF"""
        pdf_name = Path(pdf_path).stem
        output_dir = Path(output_folder) / pdf_name
        output_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                # 确保最后一个部分包含最后一页
                if sections and sections[-1]['page'] > total_pages:
                    sections[-1]['page'] = total_pages
                
                # 分割PDF
                for i in range(len(sections)):
                    start_page = sections[i]['page'] - 1
                    
                    if i < len(sections) - 1:
                        end_page = sections[i + 1]['page'] - 2
                    else:
                        end_page = total_pages - 1
                    
                    start_page = max(0, start_page)
                    end_page = min(end_page, total_pages - 1)
                    
                    if start_page > end_page:
                        continue
                    
                    pdf_writer = PyPDF2.PdfWriter()
                    
                    for page_num in range(start_page, end_page + 1):
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                    section_name = sections[i]['name']
                    safe_name = re.sub(r'[<>:"/\\|?*]', '_', section_name)
                    output_path = output_dir / f"{i+1:02d}_{safe_name}.pdf"
                    
                    with open(output_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    
                    self.log(f"    创建: {output_path.name}")
            
            return True
            
        except Exception as e:
            self.log(f"    分割失败: {str(e)}")
            return False
    
    def process_single_pdf(self, pdf_path: str, output_folder: str) -> bool:
        """处理单个PDF文件"""
        self.log(f"处理文件: {Path(pdf_path).name}")
        
        try:
            # 提取目录文本
            toc_text = self.extract_toc_text(pdf_path)
            if not toc_text:
                self.log("  警告: 未能提取目录文本，使用自动分割")
                with open(pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    total_pages = len(pdf_reader.pages)
                sections = self._auto_split_sections(total_pages)
            else:
                # 解析目录
                sections = self.parse_toc(toc_text)
                self.log(f"  解析到 {len(sections)} 个章节")
                
                # 获取总页数
                with open(pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    total_pages = len(pdf_reader.pages)
                
                # 验证和调整章节
                sections = self.validate_and_adjust_sections(sections, total_pages)
            
            # 分割PDF
            success = self.split_pdf_by_sections(pdf_path, sections, output_folder)
            
            if success:
                self.log(f"  完成!")
            else:
                self.log(f"  处理失败!")
            
            return success
            
        except Exception as e:
            self.log(f"  处理失败: {str(e)}")
            return False

class PDFSplitterGUI:
if __name__ == "__main__":
    main()