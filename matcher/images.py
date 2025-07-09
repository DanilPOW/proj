import fitz
import math

# Константы для анализа
LEFT_MARGIN = 85.04  # 3 см в пунктах
RIGHT_MARGIN = 56.69  # 2 см в пунктах
MIN_FIGURE_WIDTH = 28.35  # 1 см в пунктах
MIN_FIGURE_HEIGHT = 28.35  # 1 см в пунктах
MIN_EMPTY_LINE_DISTANCE = 12  # минимальное расстояние для пустой строки
TOLERANCE_PT = 5  # допуск в пунктах
CM_TO_PT = 28.35  # коэффициент перевода см в пункты

def analyze_pdf_images(pdf_path):
    """ПОЛНЫЙ анализ PDF с восстановленным алгоритмом поиска изображений"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Невозможно открыть PDF файл: {str(e)}")
    
    results = []
    total_pages = len(doc)
    
    # Информация о документе
    print(f"Страниц в документе: {len(doc)}")
    if hasattr(doc, 'metadata') and doc.metadata:
        print(f"Название документа: {doc.metadata.get('title', 'Н/Д')}")
    
    # Анализ каждой страницы
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            print(f"\nАнализ страницы {page_num + 1}")
            
            # Получение размеров страницы
            page_width = page.rect.width
            page_height = page.rect.height
            content_center = (LEFT_MARGIN + (page_width - RIGHT_MARGIN)) / 2
            
            print(f"Размер страницы: {page_width:.1f} x {page_height:.1f} пт")
            
            page_results = []
            
            # 1. Анализ растровых изображений
            try:
                img_list = page.get_images(full=True)
                print(f"Найдено {len(img_list)} ссылок на растровые изображения")
                
                for img_index, img in enumerate(img_list):
                    try:
                        xref = img[0]
                        print(f"Обработка изображения {img_index + 1}: xref={xref}")
                        
                        bbox = None
                        
                        # Получение bbox (используем ВСЕ методы)
                        try:
                            image_rects = page.get_image_rects()
                            for rect_info in image_rects:
                                if len(rect_info) > 1 and hasattr(rect_info[1], 'get') and rect_info[1].get('xref') == xref:
                                    bbox = rect_info[0]
                                    print(f"  Bbox (метод 1): {bbox}")
                                    break
                        except Exception as e:
                            print(f"  Ошибка метода 1: {str(e)}")
                        
                        if not bbox or bbox.is_empty:
                            try:
                                blocks = page.get_text("dict")
                                for block in blocks.get("blocks", []):
                                    if block.get("type") == 1:
                                        img_bbox = fitz.Rect(block["bbox"])
                                        if img_bbox.width >= MIN_FIGURE_WIDTH or img_bbox.height >= MIN_FIGURE_HEIGHT:
                                            bbox = img_bbox
                                            print(f"  Bbox (метод 2): {bbox}")
                                            break
                            except Exception as e:
                                print(f"  Ошибка метода 2: {str(e)}")
                        
                        if not bbox or bbox.is_empty:
                            try:
                                img_dict = doc.extract_image(xref)
                                img_width = img_dict.get("width", 100)
                                img_height = img_dict.get("height", 100)
                                
                                x0 = (page_width - img_width * 0.75) / 2
                                y0 = (page_height - img_height * 0.75) / 2
                                x1 = x0 + img_width * 0.75
                                y1 = y0 + img_height * 0.75
                                
                                bbox = fitz.Rect(x0, y0, x1, y1)
                                print(f"  Bbox (метод 3 - приблизительный): {bbox}")
                            except Exception as e:
                                print(f"  Ошибка метода 3: {str(e)}")
                        
                        if not bbox or bbox.is_empty:
                            print(f"  Не удалось получить bbox для изображения {xref}")
                            continue
                        
                        # Фильтрация по размеру
                        if (bbox.width < MIN_FIGURE_WIDTH and bbox.height < MIN_FIGURE_HEIGHT):
                            print(f"  Пропуск малого изображения {xref}: {bbox.width:.1f}x{bbox.height:.1f}")
                            continue
                        
                        # Получение свойств изображения
                        try:
                            img_dict = doc.extract_image(xref)
                            img_ext = img_dict.get("ext", "png")
                            img_width = img_dict.get("width", int(bbox.width))
                            img_height = img_dict.get("height", int(bbox.height))
                            safe_name = f"image_page{page_num+1}_xref{xref}.{img_ext}"
                        except Exception as e:
                            print(f"  Предупреждение при извлечении информации: {str(e)}")
                            safe_name = f"image_page{page_num+1}_xref{xref}.png"
                            img_width = int(bbox.width)
                            img_height = int(bbox.height)
                        
                        # Анализ позиции и полей
                        analysis = analyze_element_position(bbox, page_width, content_center)
                        
                        # НОВАЯ ПРОВЕРКА: пустая строка перед изображением
                        has_empty_line, distance, empty_line_status = check_empty_line_before_image(page, bbox)
                        
                        # Обновляем критерий соответствия ГОСТ
                        gost_compliant = analysis["margins_ok"] and analysis["is_centered"] and has_empty_line
                        
                        page_results.append({
                            "page": page_num + 1,
                            "type": "raster",
                            "bbox": [bbox.x0, bbox.y0, bbox.x1, bbox.y1],
                            "width_pt": bbox.width,
                            "height_pt": bbox.height,
                            "width_cm": bbox.width / CM_TO_PT,
                            "height_cm": bbox.height / CM_TO_PT,
                            "xref": xref,
                            "filename": safe_name,
                            "img_width": img_width,
                            "img_height": img_height,
                            "has_empty_line": has_empty_line,
                            "empty_line_distance": distance,
                            "empty_line_status": empty_line_status,
                            "gost_compliant": gost_compliant,
                            **analysis
                        })
                        
                        print(f"  ✓ Добавлено растровое изображение {xref}: {bbox.width:.1f}x{bbox.height:.1f} пт")
                        
                    except Exception as e:
                        print(f"Ошибка обработки растрового изображения {img_index}: {str(e)}")
                        continue
            
            except Exception as e:
                print(f"Ошибка получения растровых изображений: {str(e)}")
            
            # 2. Анализ векторной графики (ПОЛНЫЙ алгоритм)
            try:
                drawings = page.get_drawings()
                print(f"Найдено {len(drawings)} элементов векторной графики")
                
                significant_rects = []
                for drawing in drawings:
                    rect = drawing.get("rect")
                    if rect and rect.width > 0 and rect.height > 0:
                        if rect.width > 20 and rect.height > 20:
                            aspect_ratio = max(rect.width, rect.height) / min(rect.width, rect.height)
                            if aspect_ratio < 10:
                                significant_rects.append(rect)
                
                merged_drawings = merge_nearby_rectangles(significant_rects, merge_distance=30)
                
                final_drawings = []
                for rect in merged_drawings:
                    if (rect.width >= MIN_FIGURE_WIDTH * 2 and
                        rect.height >= MIN_FIGURE_HEIGHT * 2):
                        final_drawings.append(rect)
                
                print(f"После фильтрации: {len(final_drawings)} значимых векторных объектов")
                
                for i, rect in enumerate(final_drawings):
                    analysis = analyze_element_position(rect, page_width, content_center)
                    
                    # Проверка пустой строки для векторной графики
                    has_empty_line, distance, empty_line_status = check_empty_line_before_image(page, rect)
                    
                    gost_compliant = analysis["margins_ok"] and analysis["is_centered"] and has_empty_line
                    
                    page_results.append({
                        "page": page_num + 1,
                        "type": "vector",
                        "bbox": [rect.x0, rect.y0, rect.x1, rect.y1],
                        "width_pt": rect.width,
                        "height_pt": rect.height,
                        "width_cm": rect.width / CM_TO_PT,
                        "height_cm": rect.height / CM_TO_PT,
                        "xref": None,
                        "filename": f"vector_page{page_num+1}_item{i+1}",
                        "img_width": int(rect.width),
                        "img_height": int(rect.height),
                        "has_empty_line": has_empty_line,
                        "empty_line_distance": distance,
                        "empty_line_status": empty_line_status,
                        "gost_compliant": gost_compliant,
                        **analysis
                    })
                    
                    print(f"  ✓ Добавлена векторная графика: {rect.width:.1f}x{rect.height:.1f} пт")
            
            except Exception as e:
                print(f"Ошибка получения векторной графики: {str(e)}")
            
            # 3. Поиск встроенных изображений (ПОЛНЫЙ алгоритм)
            try:
                text_dict = page.get_text("dict")
                image_blocks = []
                
                for block in text_dict.get("blocks", []):
                    if block.get("type") == 1:
                        bbox = fitz.Rect(block["bbox"])
                        if (bbox.width >= MIN_FIGURE_WIDTH and
                             bbox.height >= MIN_FIGURE_HEIGHT):
                            
                            # КРИТИЧЕСКИ ВАЖНАЯ проверка дубликатов
                            is_duplicate = False
                            for existing in page_results:
                                existing_bbox = fitz.Rect(existing["bbox"])
                                if (abs(existing_bbox.x0 - bbox.x0) < 10 and
                                     abs(existing_bbox.y0 - bbox.y0) < 10):
                                    is_duplicate = True
                                    break
                            
                            if not is_duplicate:
                                image_blocks.append(bbox)
                
                print(f"Найдено {len(image_blocks)} дополнительных изображений через анализ содержимого")
                
                for i, bbox in enumerate(image_blocks):
                    analysis = analyze_element_position(bbox, page_width, content_center)
                    
                    has_empty_line, distance, empty_line_status = check_empty_line_before_image(page, bbox)
                    
                    gost_compliant = analysis["margins_ok"] and analysis["is_centered"] and has_empty_line
                    
                    page_results.append({
                        "page": page_num + 1,
                        "type": "embedded",
                        "bbox": [bbox.x0, bbox.y0, bbox.x1, bbox.y1],
                        "width_pt": bbox.width,
                        "height_pt": bbox.height,
                        "width_cm": bbox.width / CM_TO_PT,
                        "height_cm": bbox.height / CM_TO_PT,
                        "xref": None,
                        "filename": f"embedded_page{page_num+1}_item{i+1}",
                        "img_width": int(bbox.width),
                        "img_height": int(bbox.height),
                        "has_empty_line": has_empty_line,
                        "empty_line_distance": distance,
                        "empty_line_status": empty_line_status,
                        "gost_compliant": gost_compliant,
                        **analysis
                    })
                    
                    print(f"  ✓ Добавлено встроенное изображение: {bbox.width:.1f}x{bbox.height:.1f} пт")
            
            except Exception as e:
                print(f"Ошибка поиска встроенных изображений: {str(e)}")
            
            results.extend(page_results)
            print(f"Страница {page_num + 1}: Найдено {len(page_results)} объектов")
            
        except Exception as e:
            print(f"Ошибка обработки страницы {page_num + 1}: {str(e)}")
            continue
    
    doc.close()
    print(f"\nИТОГО: Найдено {len(results)} изображений в документе")
    return results

def merge_nearby_rectangles(rectangles, merge_distance=10):
    """Улучшенное объединение близких прямоугольников"""
    if not rectangles:
        return []
    
    rectangles = sorted(rectangles, key=lambda r: (r.y0, r.x0))
    
    merged = []
    used = set()
    
    for i, rect1 in enumerate(rectangles):
        if i in used:
            continue
        
        merged_rect = fitz.Rect(rect1)
        used.add(i)
        
        changed = True
        while changed:
            changed = False
            for j, rect2 in enumerate(rectangles):
                if j in used:
                    continue
                
                distance_x = min(abs(merged_rect.x0 - rect2.x1), abs(merged_rect.x1 - rect2.x0))
                distance_y = min(abs(merged_rect.y0 - rect2.y1), abs(merged_rect.y1 - rect2.y0))
                
                if (distance_x <= merge_distance and
                     (merged_rect.y0 <= rect2.y1 and merged_rect.y1 >= rect2.y0)) or \
                   (distance_y <= merge_distance and
                     (merged_rect.x0 <= rect2.x1 and merged_rect.x1 >= rect2.x0)) or \
                   merged_rect.intersects(rect2):
                    
                    merged_rect = merged_rect | rect2
                    used.add(j)
                    changed = True
        
        merged.append(merged_rect)
    
    return merged

def analyze_element_position(bbox, page_width, content_center):
    """Анализ позиции элемента - ТОЛЬКО ПРОВЕРКА ПОЛЕЙ"""
    center_x = (bbox.x0 + bbox.x1) / 2
    center_y = (bbox.y0 + bbox.y1) / 2
    
    margins_ok = True
    margin_errors = []
    
    if bbox.x0 < LEFT_MARGIN - TOLERANCE_PT:
        margin_errors.append("Нарушение левого поля")
        margins_ok = False
    
    if bbox.x1 > (page_width - RIGHT_MARGIN) + TOLERANCE_PT:
        margin_errors.append("Нарушение правого поля")
        margins_ok = False
    
    is_centered = abs(center_x - content_center) <= TOLERANCE_PT
    
    return {
        "center_x": center_x,
        "center_y": center_y,
        "is_centered": is_centered,
        "margins_ok": margins_ok,
        "margin_errors": margin_errors
    }

def check_empty_line_before_image(page, image_bbox):
    """Проверка пустой строки перед изображением"""
    try:
        text_dict = page.get_text("dict")
        closest_text_bottom = None
        min_distance = float('inf')
        
        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:
                block_bbox = fitz.Rect(block["bbox"])
                if block_bbox.y1 <= image_bbox.y0:
                    distance = image_bbox.y0 - block_bbox.y1
                    if distance < min_distance:
                        min_distance = distance
                        closest_text_bottom = block_bbox.y1
        
        if closest_text_bottom is None:
            return True, min_distance, "Нет текста выше изображения"
        
        has_empty_line = min_distance >= MIN_EMPTY_LINE_DISTANCE
        status = "Есть пустая строка" if has_empty_line else "Нет пустой строки"
        
        return has_empty_line, min_distance, status
        
    except Exception as e:
        return False, 0, "Ошибка проверки"
