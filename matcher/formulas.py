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

def group_nearby_spans(math_spans, vertical_tolerance=15, horizontal_gap=50):
    """
    УЛУЧШЕННАЯ группировка близких математических спанов в единые формулы
    Увеличены допуски для лучшего объединения формул
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
            # УЛУЧШЕННАЯ проверка вертикального пересечения
            # Проверяем, находятся ли спаны на одной строке или очень близко
            vertical_overlap = (
                span['bbox'].y0 <= group_span['bbox'].y1 + vertical_tolerance and
                span['bbox'].y1 >= group_span['bbox'].y0 - vertical_tolerance
            )
            
            # УЛУЧШЕННАЯ проверка горизонтальной близости
            # Учитываем как расстояние, так и возможное пересечение
            horizontal_distance = float('inf')
            
            # Расстояние между правым краем группы и левым краем нового спана
            dist1 = abs(span['bbox'].x0 - group_span['bbox'].x1)
            # Расстояние между левым краем группы и правым краем нового спана  
            dist2 = abs(group_span['bbox'].x0 - span['bbox'].x1)
            
            horizontal_distance = min(dist1, dist2)
            
            # Проверяем пересечение
            horizontal_intersects = span['bbox'].intersects(group_span['bbox'])
            
            horizontal_close = (
                horizontal_distance <= horizontal_gap or
                horizontal_intersects
            )
            
            # ДОПОЛНИТЕЛЬНАЯ проверка для математических выражений
            # Если спаны содержат математические символы и находятся близко
            math_symbols = ['=', '+', '-', '×', '÷', '/', '^', '²', '³', '∑', '∫', '∂']
            span_has_math = any(symbol in span['text'] for symbol in math_symbols)
            group_has_math = any(symbol in group_span['text'] for symbol in math_symbols)
            
            if span_has_math or group_has_math:
                # Для математических выражений используем более щедрые допуски
                if vertical_overlap and horizontal_distance <= horizontal_gap * 1.5:
                    can_merge = True
                    break
            elif vertical_overlap and horizontal_close:
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
    
    # ПОСТОБРАБОТКА: объединяем группы, которые могут быть частями одной формулы
    final_groups = []
    i = 0
    while i < len(groups):
        current_group = groups[i]
        
        # Проверяем, можно ли объединить с следующей группой
        if i + 1 < len(groups):
            next_group = groups[i + 1]
            
            # Вычисляем расстояние между группами
            current_bbox = get_group_bbox(current_group)
            next_bbox = get_group_bbox(next_group)
            
            # Если группы на одной строке и близко друг к другу
            vertical_close = abs(current_bbox.y0 - next_bbox.y0) <= vertical_tolerance
            horizontal_distance = next_bbox.x0 - current_bbox.x1
            
            if vertical_close and horizontal_distance <= horizontal_gap:
                # Объединяем группы
                current_group.extend(next_group)
                i += 2  # Пропускаем следующую группу
            else:
                i += 1
        else:
            i += 1
        
        final_groups.append(current_group)
    
    return final_groups

def get_group_bbox(span_group):
    """Вычисляет общий bbox для группы спанов"""
    if not span_group:
        return fitz.Rect()
    
    min_x0 = min(span['bbox'].x0 for span in span_group)
    min_y0 = min(span['bbox'].y0 for span in span_group)
    max_x1 = max(span['bbox'].x1 for span in span_group)
    max_y1 = max(span['bbox'].y1 for span in span_group)
    
    return fitz.Rect(min_x0, min_y0, max_x1, max_y1)

def merge_span_group(span_group):
    """
    УЛУЧШЕННОЕ объединение группы спанов в единую формулу
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
    
    # УЛУЧШЕННОЕ объединение текста
    combined_text = ""
    for i, span in enumerate(sorted_spans):
        if i > 0:
            prev_span = sorted_spans[i-1]
            
            # Вычисляем расстояние между спанами
            gap = span['bbox'].x0 - prev_span['bbox'].x1
            
            # Проверяем, нужен ли пробел
            prev_text = prev_span['text'].strip()
            curr_text = span['text'].strip()
            
            # Не добавляем пробел если:
            # - предыдущий текст заканчивается на математический символ
            # - текущий текст начинается с математического символа
            # - расстояние очень маленькое
            math_endings = ['=', '+', '-', '×', '÷', '/', '^', '(', '[']
            math_beginnings = ['=', '+', '-', '×', '÷', '/', '^', ')', ']', '²', '³']
            
            needs_space = True
            if gap <= 2:  # Очень близко
                needs_space = False
            elif prev_text and prev_text[-1] in math_endings:
                needs_space = False
            elif curr_text and curr_text[0] in math_beginnings:
                needs_space = False
            elif gap > 3 and gap <= 10:  # Средний разрыв
                needs_space = True
            elif gap > 10:  # Большой разрыв - возможно, нужен пробел
                needs_space = True
            
            if needs_space:
                combined_text += " "
        
        combined_text += span['text']
    
    # Очищаем лишние пробелы
    combined_text = ' '.join(combined_text.split())
    
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
