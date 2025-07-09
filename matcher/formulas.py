import fitz
import re

# Константы
LEFT_MARGIN = 85.04  # 3 см в пунктах
RIGHT_MARGIN = 56.69  # 2 см в пунктах
TOLERANCE_PT = 5

# Список математических шрифтов
MATH_FONTS = {
    "cambriamath",
    "cambria-math", 
    "cambriamt",
    "cmmi",
    "cmr",
    "cmsy",
    "cmex",
    "stix",
    "stixmath",
    "mathtime",
    "xits",
    "xitsmath",
    "latinmodernmath",
    "latinmodern-math",
    "texgyrepagella",
    "texgyretermes",
    "asanamath",
    "neoeuler"
}

def normalize_font_name(font_name):
    """Нормализация названия шрифта для сравнения"""
    if not font_name:
        return ""
    return font_name.lower().replace(" ", "").replace("-", "").replace("_", "")

def is_math_font(font_name):
    """Проверка, является ли шрифт математическим"""
    normalized = normalize_font_name(font_name)
    return any(math_font in normalized for math_font in MATH_FONTS)

def group_nearby_spans(math_spans, vertical_tolerance=8, horizontal_gap=30):
    """
    Группировка близких математических спанов в единые формулы
    """
    if not math_spans:
        return []
    
    # Сортируем спаны по позиции (сначала по Y, потом по X)
    sorted_spans = sorted(math_spans, key=lambda s: (s['bbox'].y0, s['bbox'].x0))
    
    groups = []
    current_group = [sorted_spans[0]]
    
    for span in sorted_spans[1:]:
        # Проверяем, можно ли добавить спан к текущей группе
        can_merge = False
        
        for group_span in current_group:
            # Проверяем вертикальное пересечение или близость
            vertical_overlap = (
                span['bbox'].y0 <= group_span['bbox'].y1 + vertical_tolerance and
                span['bbox'].y1 >= group_span['bbox'].y0 - vertical_tolerance
            )
            
            # Проверяем горизонтальную близость
            horizontal_distance = min(
                abs(span['bbox'].x0 - group_span['bbox'].x1),
                abs(group_span['bbox'].x0 - span['bbox'].x1)
            )
            
            horizontal_close = (
                horizontal_distance <= horizontal_gap or
                span['bbox'].intersects(group_span['bbox'])
            )
            
            if vertical_overlap and horizontal_close:
                can_merge = True
                break
        
        if can_merge:
            current_group.append(span)
        else:
            # Начинаем новую группу
            groups.append(current_group)
            current_group = [span]
    
    # Добавляем последнюю группу
    if current_group:
        groups.append(current_group)
    
    return groups

def merge_span_group(span_group):
    """
    Объединение группы спанов в единую формулу
    """
    if not span_group:
        return None
    
    # Вычисляем общий bbox
    min_x0 = min(span['bbox'].x0 for span in span_group)
    min_y0 = min(span['bbox'].y0 for span in span_group)
    max_x1 = max(span['bbox'].x1 for span in span_group)
    max_y1 = max(span['bbox'].y1 for span in span_group)
    
    combined_bbox = fitz.Rect(min_x0, min_y0, max_x1, max_y1)
    
    # Сортируем спаны по позиции для правильного порядка текста
    sorted_spans = sorted(span_group, key=lambda s: (s['bbox'].y0, s['bbox'].x0))
    
    # Объединяем текст
    combined_text = ""
    for i, span in enumerate(sorted_spans):
        if i > 0:
            # Добавляем пробел между спанами, если они не слишком близко
            prev_span = sorted_spans[i-1]
            gap = span['bbox'].x0 - prev_span['bbox'].x1
            if gap > 3:  # Если разрыв больше 3 пунктов
                combined_text += " "
        combined_text += span['text']
    
    # Собираем информацию о шрифтах
    fonts_used = list(set(span['font'] for span in span_group))
    
    return {
        'text': combined_text.strip(),
        'bbox': combined_bbox,
        'fonts': fonts_used,
        'span_count': len(span_group),
        'spans': span_group
    }

def find_formula_numbering(page, formula_bbox):
    """
    Поиск нумерации формулы справа от неё
    """
    try:
        # Расширяем область поиска вправо от формулы
        search_area = fitz.Rect(
            formula_bbox.x1,
            formula_bbox.y0 - 5,
            formula_bbox.x1 + 100,
            formula_bbox.y1 + 5
        )
        
        # Получаем текст в области поиска
        text_in_area = page.get_text("text", clip=search_area).strip()
        
        # Ищем паттерны нумерации: (1), (2), [1], [2] и т.д.
        numbering_patterns = [
            r'$$\s*(\d+)\s*$$',  # (1), (2)
            r'$$\s*(\d+)\s*$$',  # [1], [2]
            r'$$\s*(\d+\.\d+)\s*$$',  # (1.1), (2.3)
        ]
        
        for pattern in numbering_patterns:
            match = re.search(pattern, text_in_area)
            if match:
                return True, match.group(0)
        
        return False, ""
        
    except Exception as e:
        print(f"Ошибка поиска нумерации формулы: {e}")
        return False, ""

def analyze_pdf_formulas(pdf_path, existing_images):
    """
    ПРИОРИТЕТНЫЙ анализ формул в PDF по математическому шрифту
    Теперь выполняется ПЕРВЫМ, без блокировки изображениями
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise Exception(f"Невозможно открыть PDF файл: {str(e)}")
    
    results = []
    
    print(f"🚀 ПРИОРИТЕТНЫЙ поиск формул по математическому шрифту (БЕЗ блокировки изображениями)")
    
    for page_num in range(len(doc)):
        try:
            page = doc.load_page(page_num)
            text_dict = page.get_text("dict")
            
            print(f"Анализ формул на странице {page_num + 1}")
            
            # Собираем все шрифты на странице
            all_fonts = set()
            math_spans = []
            
            for block in text_dict.get("blocks", []):
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            font_name = span.get("font", "")
                            text = span.get("text", "").strip()
                            
                            if font_name:
                                all_fonts.add(font_name)
                            
                            # Проверяем, является ли шрифт математическим
                            if font_name and text and is_math_font(font_name):
                                bbox = fitz.Rect(span["bbox"])
                                
                                math_spans.append({
                                    'text': text,
                                    'font': font_name,
                                    'bbox': bbox,
                                    'size': span.get("size", 12)
                                })
                                
                                print(f"  ✅ Найден математический спан: '{text}' (шрифт: {font_name})")
            
            print(f"  Все шрифты на странице: {sorted(all_fonts)}")
            print(f"  Найдено {len(math_spans)} спанов с математическим шрифтом")
            
            if not math_spans:
                continue
            
            # Группируем близкие математические спаны
            span_groups = group_nearby_spans(math_spans)
            print(f"  Сгруппировано в {len(span_groups)} формул")
            
            # Обрабатываем каждую группу
            for group_idx, span_group in enumerate(span_groups):
                try:
                    # Объединяем спаны в группе
                    formula_info = merge_span_group(span_group)
                    if not formula_info:
                        continue
                    
                    formula_text = formula_info['text']
                    formula_bbox = formula_info['bbox']
                    
                    print(f"  Формула {group_idx + 1}: '{formula_text}' (спанов: {formula_info['span_count']})")
                    
                    # БЕЗ ПРОВЕРКИ ПЕРЕСЕЧЕНИЯ С ИЗОБРАЖЕНИЯМИ!
                    # Формулы имеют приоритет
                    
                    # Анализ позиции формулы
                    page_width = page.rect.width
                    content_center = (LEFT_MARGIN + (page_width - RIGHT_MARGIN)) / 2
                    formula_center = (formula_bbox.x0 + formula_bbox.x1) / 2
                    
                    is_centered = abs(formula_center - content_center) <= TOLERANCE_PT * 4  # Увеличенный допуск
                    margins_ok = (formula_bbox.x0 >= LEFT_MARGIN - TOLERANCE_PT and
                                 formula_bbox.x1 <= page_width - RIGHT_MARGIN + TOLERANCE_PT)
                    
                    # Поиск нумерации
                    has_numbering, numbering_text = find_formula_numbering(page, formula_bbox)
                    
                    # МЯГКОЕ определение соответствия ГОСТ
                    gost_compliant = is_centered and margins_ok
                    
                    results.append({
                        "page": page_num + 1,
                        "text": formula_text,
                        "bbox": [formula_bbox.x0, formula_bbox.y0, formula_bbox.x1, formula_bbox.y1],
                        "is_centered": is_centered,
                        "margins_ok": margins_ok,
                        "has_numbering": has_numbering,
                        "numbering_text": numbering_text,
                        "fonts": formula_info['fonts'],
                        "span_count": formula_info['span_count'],
                        "type": "math_font_detected",
                        "gost_compliant": gost_compliant
                    })
                    
                    print(f"    🎯 ДОБАВЛЕНА ФОРМУЛА: '{formula_text}' - центр={is_centered}, поля={margins_ok}, номер={has_numbering}")
                    if has_numbering:
                        print(f"      📝 Нумерация: {numbering_text}")
                
                except Exception as e:
                    print(f"    Ошибка обработки группы {group_idx}: {e}")
                    continue
        
        except Exception as e:
            print(f"Ошибка анализа формул на странице {page_num + 1}: {e}")
            continue
    
    doc.close()
    print(f"🏁 ПРИОРИТЕТНЫЙ поиск формул завершен. Найдено: {len(results)} формул")
    
    # Выводим найденные формулы для проверки
    for i, formula in enumerate(results, 1):
        print(f"  📐 Формула {i}: '{formula['text']}' (стр: {formula['page']}, шрифты: {formula['fonts']})")
        if formula['has_numbering']:
            print(f"    📝 Нумерация: {formula['numbering_text']}")
    
    return results
