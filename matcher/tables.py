import fitz

# Константы
LEFT_MARGIN = 85.04  # 3 см в пунктах
RIGHT_MARGIN = 56.69  # 2 см в пунктах
TOLERANCE_PT = 5
CM_TO_PT = 28.35

def analyze_pdf_tables(pdf_path, existing_images):
    """ИСПРАВЛЕННЫЙ анализ таблиц в PDF"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Невозможно открыть PDF файл: {str(e)}")
    
    results = []
    
    # Получаем список уже найденных изображений для исключения
    image_areas = []
    for img in existing_images:
        image_areas.append(fitz.Rect(img['bbox']))
    
    print(f"Начинаем поиск таблиц. Исключаем {len(image_areas)} областей изображений")
    
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            print(f"Анализ таблиц на странице {page_num + 1}")
            
            # Метод 1: Используем встроенный поиск таблиц PyMuPDF
            try:
                tables = page.find_tables()
                print(f"  Метод 1 (find_tables): найдено {len(tables)} таблиц")
                
                for table_index, table in enumerate(tables):
                    try:
                        bbox = table.bbox
                        if isinstance(bbox, (list, tuple)):
                            bbox = fitz.Rect(bbox)
                        
                        # Проверяем минимальный размер
                        if bbox.width < 50 or bbox.height < 30:
                            print(f"    Пропуск таблицы {table_index + 1} - слишком маленькая")
                            continue
                        
                        # Проверяем, не пересекается ли с изображениями
                        is_image = False
                        for img_bbox in image_areas:
                            if bbox.intersects(img_bbox):
                                overlap_area = (bbox & img_bbox).get_area()
                                bbox_area = bbox.get_area()
                                if bbox_area > 0 and overlap_area / bbox_area > 0.5:
                                    is_image = True
                                    break
                        
                        if is_image:
                            print(f"    Пропуск таблицы {table_index + 1} - пересекается с изображением")
                            continue
                        
                        # Анализ позиции таблицы
                        page_width = page.rect.width
                        content_center = (LEFT_MARGIN + (page_width - RIGHT_MARGIN)) / 2
                        
                        table_center = (bbox.x0 + bbox.x1) / 2
                        is_centered = abs(table_center - content_center) <= TOLERANCE_PT * 2  # Увеличенный допуск
                        margins_ok = (bbox.x0 >= LEFT_MARGIN - TOLERANCE_PT and
                                     bbox.x1 <= page_width - RIGHT_MARGIN + TOLERANCE_PT)
                        
                        # Проверка наличия заголовка и нумерации
                        has_title = check_table_title(page, bbox)
                        has_numbering = check_table_numbering(page, bbox)
                        
                        # Получаем содержимое таблицы для подсчета строк
                        try:
                            table_data = table.extract()
                            row_count = len(table_data) if table_data else 0
                        except:
                            row_count = 0
                        
                        gost_compliant = is_centered and margins_ok and has_title and has_numbering
                        
                        results.append({
                            "page": page_num + 1,
                            "bbox": [bbox.x0, bbox.y0, bbox.x1, bbox.y1],
                            "width_cm": bbox.width / CM_TO_PT,
                            "height_cm": bbox.height / CM_TO_PT,
                            "is_centered": is_centered,
                            "margins_ok": margins_ok,
                            "has_title": has_title,
                            "has_numbering": has_numbering,
                            "rows": row_count,
                            "gost_compliant": gost_compliant,
                            "table_num": table_index + 1,
                            "detection_method": "find_tables"
                        })
                        
                        print(f"    ✓ Найдена таблица {table_index + 1}: {bbox.width:.1f}x{bbox.height:.1f} пт, строк: {row_count}")
                        
                    except Exception as e:
                        print(f"    Ошибка анализа таблицы {table_index}: {e}")
                        continue
            
            except Exception as e:
                print(f"  Ошибка метода find_tables: {e}")
            
        except Exception as e:
            print(f"Ошибка анализа таблиц на странице {page_num + 1}: {e}")
            continue
    
    doc.close()
    print(f"Поиск таблиц завершен. Найдено: {len(results)} таблиц")
    return results

def check_table_title(page, table_bbox):
    """Проверка наличия заголовка таблицы"""
    try:
        # Ищем текст "Таблица" в области выше таблицы
        search_area = fitz.Rect(
            table_bbox.x0 - 50,
            table_bbox.y0 - 100,
            table_bbox.x1 + 50,
            table_bbox.y0
        )
        
        text_in_area = page.get_text("text", clip=search_area).lower()
        return "таблица" in text_in_area
        
    except Exception as e:
        print(f"Ошибка проверки заголовка таблицы: {e}")
        return False

def check_table_numbering(page, table_bbox):
    """Проверка нумерации таблицы"""
    try:
        # Ищем номер таблицы в области выше таблицы
        search_area = fitz.Rect(
            table_bbox.x0 - 50,
            table_bbox.y0 - 100,
            table_bbox.x1 + 50,
            table_bbox.y0
        )
        
        text_in_area = page.get_text("text", clip=search_area)
        # Ищем паттерн "Таблица N" или просто цифры
        import re
        return bool(re.search(r'таблица\s+\d+|^\d+', text_in_area, re.IGNORECASE))
        
    except Exception as e:
        print(f"Ошибка проверки нумерации таблицы: {e}")
        return False
