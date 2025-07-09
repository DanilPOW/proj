import fitz
import re

def analyze_page_numbering(pdf_path):
    """Анализ нумерации страниц в PDF - МЯГКАЯ ПРОВЕРКА ЦЕНТРИРОВАНИЯ"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Невозможно открыть PDF файл: {str(e)}")
    
    results = []
    
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            page_rect = page.rect
            
            # Поиск номера страницы в нижней части
            bottom_area = fitz.Rect(0, page_rect.height - 50, page_rect.width, page_rect.height)
            bottom_text = page.get_text("text", clip=bottom_area)
            
            # Поиск числа в нижней части страницы
            page_number_found = False
            expected_number = page_num + 1
            is_correct_number = False
            is_centered = False
            
            # Ищем номер страницы
            numbers_in_bottom = re.findall(r'\b\d+\b', bottom_text)
            
            if numbers_in_bottom:
                page_number_found = True
                # Проверяем, есть ли правильный номер
                if str(expected_number) in numbers_in_bottom:
                    is_correct_number = True
                    
                    # МЯГКАЯ ПРОВЕРКА центрирования (увеличенный допуск)
                    text_blocks = page.get_text("dict", clip=bottom_area)
                    for block in text_blocks.get("blocks", []):
                        if "lines" in block:
                            for line in block["lines"]:
                                for span in line["spans"]:
                                    if any(num in span["text"] for num in numbers_in_bottom):
                                        span_center = (span["bbox"][0] + span["bbox"][2]) / 2
                                        page_center = page_rect.width / 2
                                        # Увеличенный допуск для центрирования (30 пт вместо 5)
                                        is_centered = abs(span_center - page_center) <= 30
            
            # Для первой страницы номер не должен быть
            if page_num == 0:
                gost_compliant = not page_number_found
                expected_status = "Не должно быть номера"
            else:
                gost_compliant = page_number_found and is_correct_number and is_centered
                expected_status = f"Номер {expected_number}, по центру"
            
            results.append({
                "page": page_num + 1,
                "has_number": page_number_found,
                "expected_number": expected_number,
                "is_correct_number": is_correct_number,
                "is_centered": is_centered,
                "expected_status": expected_status,
                "found_numbers": numbers_in_bottom,
                "gost_compliant": gost_compliant
            })
            
        except Exception as e:
            print(f"Ошибка анализа нумерации страницы {page_num + 1}: {e}")
            continue
    
    doc.close()
    return results
