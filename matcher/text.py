import fitz

# Константы
LEFT_MARGIN = 85.04  # 3 см в пунктах
RIGHT_MARGIN = 56.69  # 2 см в пунктах
TOLERANCE_PT = 5

def analyze_pdf_text(pdf_path):
    """Анализ текста в PDF - ПРОВЕРКА ПОЛЕЙ И РАЗМЕРА ШРИФТА (14pt вместо 12pt)"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Невозможно открыть PDF файл: {str(e)}")
    
    results = []
    
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            text_dict = page.get_text("dict")
            
            for block in text_dict.get("blocks", []):
                if "lines" in block:
                    for line_num, line in enumerate(block["lines"]):
                        for span_num, span in enumerate(line["spans"]):
                            text = span["text"].strip()
                            if not text:
                                continue
                            
                            # Проверка размера шрифта (теперь 14pt вместо 12pt)
                            font_size = span.get("size", 14)
                            if abs(font_size - 14) > 1:  # допуск 1 пт для 14pt
                                results.append({
                                    "page": page_num + 1,
                                    "issue": "font_size",
                                    "font": span.get("font", "Unknown"),
                                    "size": f"{font_size:.1f}pt",
                                    "margins": "N/A",
                                    "text_preview": text[:50] + "..." if len(text) > 50 else text,
                                    "gost_compliant": False
                                })
                            
                            # Проверка шрифта (добавлен Cambria Math в список разрешенных)
                            font_name = span.get("font", "")
                            if font_name and not any(allowed in font_name.lower()
                                                    for allowed in ["times", "arial", "calibri", "cambria"]):
                                results.append({
                                    "page": page_num + 1,
                                    "issue": "font_family",
                                    "font": font_name,
                                    "size": f"{font_size:.1f}pt",
                                    "margins": "N/A",
                                    "text_preview": text[:50] + "..." if len(text) > 50 else text,
                                    "gost_compliant": False
                                })
                            
                            # ТОЛЬКО ПРОВЕРКА ПОЛЕЙ (убрана проверка выравнивания)
                            bbox = fitz.Rect(span["bbox"])
                            page_width = page.rect.width
                            
                            left_margin_check = bbox.x0 < LEFT_MARGIN - TOLERANCE_PT
                            right_margin_check = bbox.x1 > page_width - RIGHT_MARGIN + TOLERANCE_PT
                            
                            if left_margin_check or right_margin_check:
                                margin_issue = []
                                if left_margin_check:
                                    margin_issue.append("левое поле")
                                if right_margin_check:
                                    margin_issue.append("правое поле")
                                
                                results.append({
                                    "page": page_num + 1,
                                    "issue": "margins",
                                    "font": font_name,
                                    "size": f"{font_size:.1f}pt",
                                    "margins": f"Нарушение: {', '.join(margin_issue)}",
                                    "text_preview": text[:50] + "..." if len(text) > 50 else text,
                                    "gost_compliant": False
                                })
        
        except Exception as e:
            print(f"Ошибка анализа текста на странице {page_num + 1}: {e}")
            continue
    
    doc.close()
    return results
