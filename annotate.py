import fitz
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import tempfile
import os
import zipfile
import xml.etree.ElementTree as ET
import re

def remove_existing_comments(docx_path, output_path):
    """
    БЕЗОПАСНОЕ удаление всех существующих примечаний из DOCX файла
    
    Args:
        docx_path: путь к исходному DOCX файлу
        output_path: путь для сохранения очищенного файла
    """
    try:
        print(f"Удаление примечаний из {docx_path}")
        
        # Метод 1: Простое удаление через python-docx (БЕЗОПАСНЫЙ)
        try:
            doc = Document(docx_path)
            
            # Удаляем примечания более безопасным способом
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    # Удаляем элементы комментариев без использования xpath
                    run_element = run._element
                    
                    # Находим и удаляем элементы комментариев
                    comment_elements = []
                    for child in run_element:
                        if child.tag.endswith('}commentRangeStart') or \
                           child.tag.endswith('}commentRangeEnd') or \
                           child.tag.endswith('}commentReference'):
                            comment_elements.append(child)
                    
                    # Удаляем найденные элементы
                    for elem in comment_elements:
                        run_element.remove(elem)
            
            # Удаляем примечания из таблиц
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run_element = run._element
                                comment_elements = []
                                for child in run_element:
                                    if child.tag.endswith('}commentRangeStart') or \
                                       child.tag.endswith('}commentRangeEnd') or \
                                       child.tag.endswith('}commentReference'):
                                        comment_elements.append(child)
                                
                                for elem in comment_elements:
                                    run_element.remove(elem)
            
            doc.save(output_path)
            print("✓ Метод 1 (безопасный python-docx) выполнен")
            
            # Проверяем, что файл не поврежден
            try:
                test_doc = Document(output_path)
                print("✓ Файл прошел проверку целостности")
                return True
            except Exception as e:
                print(f"⚠️ Файл поврежден после метода 1: {e}")
                # Переходим к методу 2
                
        except Exception as e:
            print(f"Ошибка метода 1: {e}")
        
        # Метод 2: Более осторожная работа с ZIP (ИСПРАВЛЕННЫЙ)
        try:
            print("Применяем метод 2 (осторожная работа с ZIP)")
            
            # Создаем временную копию
            temp_path = output_path + '.temp'
            import shutil
            shutil.copy2(docx_path, temp_path)  # Копируем ИСХОДНЫЙ файл
            
            # Открываем как ZIP архив
            with zipfile.ZipFile(temp_path, 'r') as zip_read:
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
                    for item in zip_read.infolist():
                        data = zip_read.read(item.filename)
                        
                        # Пропускаем файл comments.xml
                        if item.filename == 'word/comments.xml':
                            print("  Удален word/comments.xml")
                            continue
                        
                        # Обрабатываем document.xml БОЛЕЕ ОСТОРОЖНО
                        if item.filename == 'word/document.xml':
                            try:
                                content = data.decode('utf-8')
                                
                                # Используем регулярные выражения для удаления комментариев
                                # Удаляем commentRangeStart
                                content = re.sub(r'<w:commentRangeStart[^>]*?/>', '', content)
                                # Удаляем commentRangeEnd
                                content = re.sub(r'<w:commentRangeEnd[^>]*?/>', '', content)
                                # Удаляем commentReference
                                content = re.sub(r'<w:commentReference[^>]*?/>', '', content)
                                
                                data = content.encode('utf-8')
                                print("  Обработан word/document.xml (регулярные выражения)")
                                
                            except Exception as e:
                                print(f"  Ошибка обработки document.xml: {e}")
                        
                        # Обрабатываем [Content_Types].xml
                        elif item.filename == '[Content_Types].xml':
                            try:
                                content = data.decode('utf-8')
                                # Удаляем строки с comments
                                lines = content.split('\n')
                                filtered_lines = [line for line in lines if 'comments' not in line.lower()]
                                data = '\n'.join(filtered_lines).encode('utf-8')
                                print("  Обработан [Content_Types].xml")
                            except Exception as e:
                                print(f"  Ошибка обработки [Content_Types].xml: {e}")
                        
                        # Обрабатываем word/_rels/document.xml.rels
                        elif item.filename == 'word/_rels/document.xml.rels':
                            try:
                                content = data.decode('utf-8')
                                lines = content.split('\n')
                                filtered_lines = [line for line in lines if 'comments' not in line.lower()]
                                data = '\n'.join(filtered_lines).encode('utf-8')
                                print("  Обработан word/_rels/document.xml.rels")
                            except Exception as e:
                                print(f"  Ошибка обработки document.xml.rels: {e}")
                        
                        zip_write.writestr(item, data)
            
            # Удаляем временный файл
            os.unlink(temp_path)
            
            # Проверяем целостность файла
            try:
                test_doc = Document(output_path)
                print("✓ Метод 2 (ZIP) выполнен и файл прошел проверку")
                return True
            except Exception as e:
                print(f"⚠️ Файл поврежден после метода 2: {e}")
                
        except Exception as e:
            print(f"Ошибка метода 2: {e}")
        
        # Метод 3: Если все не удалось, просто копируем исходный файл
        print("Применяем метод 3 (копирование исходного файла)")
        import shutil
        shutil.copy2(docx_path, output_path)
        print("⚠️ Примечания не удалены, но файл сохранен")
        return False
        
    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА при удалении примечаний: {e}")
        # В крайнем случае просто копируем файл
        try:
            import shutil
            shutil.copy2(docx_path, output_path)
            return False
        except:
            return False

def find_docx_images_with_positions(docx_path):
    """
    Находит все изображения в DOCX с их позициями и контекстом
    
    Returns:
        list: список словарей с информацией об изображениях
    """
    try:
        doc = Document(docx_path)
        images_info = []
        
        # Ищем изображения в параграфах
        for para_idx, paragraph in enumerate(doc.paragraphs):
            for run_idx, run in enumerate(paragraph.runs):
                # Проверяем наличие изображений в run
                if run._element.xpath('.//pic:pic'):
                    # Получаем контекст - текст до и после
                    context_before = ""
                    context_after = ""
                    
                    # Текст из предыдущих параграфов
                    for i in range(max(0, para_idx - 2), para_idx):
                        context_before += doc.paragraphs[i].text + " "
                    
                    # Текст из текущего параграфа до изображения
                    for i in range(run_idx):
                        context_before += paragraph.runs[i].text + " "
                    
                    # Текст из текущего параграфа после изображения
                    for i in range(run_idx + 1, len(paragraph.runs)):
                        context_after += paragraph.runs[i].text + " "
                    
                    # Текст из следующих параграфов
                    for i in range(para_idx + 1, min(len(doc.paragraphs), para_idx + 3)):
                        context_after += doc.paragraphs[i].text + " "
                    
                    images_info.append({
                        'paragraph_idx': para_idx,
                        'run_idx': run_idx,
                        'context_before': context_before.strip()[-100:],  # Последние 100 символов
                        'context_after': context_after.strip()[:100],    # Первые 100 символов
                        'paragraph_text': paragraph.text,
                        'position_in_doc': para_idx  # Позиция в документе
                    })
        
        print(f"Найдено {len(images_info)} изображений в DOCX")
        for i, img in enumerate(images_info):
            print(f"  Изображение {i+1}: параграф {img['paragraph_idx']}, контекст: '{img['context_before'][-50:]}' ... '{img['context_after'][:50]}'")
        
        return images_info
        
    except Exception as e:
        print(f"Ошибка поиска изображений в DOCX: {e}")
        return []

def find_docx_formulas_with_positions(docx_path):
    """
    НОВАЯ ФУНКЦИЯ: Находит все формулы в DOCX с их позициями
    
    Returns:
        list: список словарей с информацией о формулах
    """
    try:
        doc = Document(docx_path)
        formulas_info = []
        
        # Ищем формулы в параграфах
        for para_idx, paragraph in enumerate(doc.paragraphs):
            para_text = paragraph.text.strip()
            
            # Проверяем, содержит ли параграф математические символы или формулы
            math_indicators = ['=', '∑', '∫', '∂', '∆', '∇', '±', '×', '÷', '≤', '≥', '≠', '≈', '∞']
            has_math = any(indicator in para_text for indicator in math_indicators)
            
            # Проверяем наличие математических объектов Word
            has_math_object = False
            for run in paragraph.runs:
                # Проверяем наличие математических объектов в run
                if run._element.xpath('.//m:oMath') or run._element.xpath('.//m:oMathPara'):
                    has_math_object = True
                    break
            
            # Проверяем центрированность параграфа (часто формулы центрированы)
            is_centered = False
            if paragraph.paragraph_format.alignment is not None:
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                is_centered = paragraph.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
            
            # Считаем формулой если есть математические символы или объекты
            if has_math or has_math_object:
                # Получаем контекст
                context_before = ""
                context_after = ""
                
                # Текст из предыдущих параграфов
                for i in range(max(0, para_idx - 2), para_idx):
                    context_before += doc.paragraphs[i].text + " "
                
                # Текст из следующих параграфов
                for i in range(para_idx + 1, min(len(doc.paragraphs), para_idx + 3)):
                    context_after += doc.paragraphs[i].text + " "
                
                formulas_info.append({
                    'paragraph_idx': para_idx,
                    'text': para_text,
                    'context_before': context_before.strip()[-100:],
                    'context_after': context_after.strip()[:100],
                    'has_math_symbols': has_math,
                    'has_math_object': has_math_object,
                    'is_centered': is_centered,
                    'position_in_doc': para_idx
                })
        
        print(f"Найдено {len(formulas_info)} потенциальных формул в DOCX")
        for i, formula in enumerate(formulas_info):
            print(f"  Формула {i+1}: параграф {formula['paragraph_idx']}, текст: '{formula['text'][:50]}...', центр: {formula['is_centered']}")
        
        return formulas_info
        
    except Exception as e:
        print(f"Ошибка поиска формул в DOCX: {e}")
        return []

def match_pdf_images_to_docx(pdf_images, docx_images):
    """
    УЛУЧШЕННОЕ сопоставление изображений PDF с DOCX
    
    Args:
        pdf_images: список изображений из PDF анализа
        docx_images: список изображений из DOCX
    
    Returns:
        dict: сопоставление {pdf_image_index: docx_image_index}
    """
    matching = {}
    
    if not pdf_images or not docx_images:
        return matching
    
    print(f"Сопоставление: {len(pdf_images)} изображений PDF с {len(docx_images)} изображениями DOCX")
    
    # Группируем PDF изображения по страницам
    pdf_by_page = {}
    for i, pdf_img in enumerate(pdf_images):
        page = pdf_img.get('page', 1)
        if page not in pdf_by_page:
            pdf_by_page[page] = []
        pdf_by_page[page].append((i, pdf_img))
    
    # Простое сопоставление: первое изображение на странице -> первое изображение в DOCX
    docx_idx = 0
    for page in sorted(pdf_by_page.keys()):
        page_images = pdf_by_page[page]
        
        for pdf_idx, pdf_img in page_images:
            if docx_idx < len(docx_images):
                matching[pdf_idx] = docx_idx
                print(f"  PDF изображение {pdf_idx} (стр. {page}) -> DOCX изображение {docx_idx}")
                docx_idx += 1
            else:
                print(f"  PDF изображение {pdf_idx} (стр. {page}) -> НЕТ СООТВЕТСТВИЯ в DOCX")
    
    return matching

def match_pdf_formulas_to_docx(pdf_formulas, docx_formulas):
    """
    НОВАЯ ФУНКЦИЯ: Сопоставление формул PDF с DOCX
    
    Args:
        pdf_formulas: список формул из PDF анализа
        docx_formulas: список формул из DOCX
    
    Returns:
        dict: сопоставление {pdf_formula_index: docx_formula_index}
    """
    matching = {}
    
    if not pdf_formulas or not docx_formulas:
        print(f"Сопоставление формул: PDF={len(pdf_formulas)}, DOCX={len(docx_formulas)} - недостаточно данных")
        return matching
    
    print(f"Сопоставление: {len(pdf_formulas)} формул PDF с {len(docx_formulas)} формулами DOCX")
    
    # Простое сопоставление по порядку появления
    for i, pdf_formula in enumerate(pdf_formulas):
        if i < len(docx_formulas):
            matching[i] = i
            print(f"  PDF формула {i} (стр. {pdf_formula.get('page', '?')}, '{pdf_formula.get('text', '')[:30]}...') -> DOCX формула {i}")
        else:
            print(f"  PDF формула {i} (стр. {pdf_formula.get('page', '?')}) -> НЕТ СООТВЕТСТВИЯ в DOCX")
    
    return matching

def try_commenting(doc, para_index, comment_text, method="first"):
    """
    Добавляет примечание к параграфу различными методами
    
    Args:
        doc: объект Document
        para_index: индекс параграфа
        comment_text: текст примечания
        method: метод добавления ("first", "last", "all")
    
    Returns:
        bool: успешность операции
    """
    try:
        if para_index >= len(doc.paragraphs):
            return False
        
        para = doc.paragraphs[para_index]
        runs = para.runs
        
        if not runs:
            # Создаём пустой run если его нет
            runs = [para.add_run(" ")]
        
        if method == "first":
            target_runs = runs[0]
        elif method == "last":
            target_runs = runs[-1]
        else:  # "all"
            target_runs = runs
        
        # Добавляем примечание
        comment = doc.add_comment(runs=target_runs, text=comment_text, author="ГОСТ Анализатор", initials="ГА")
        print(f"Добавлено примечание методом {method}: {comment_text[:50]}...")
        return True
        
    except Exception as e:
        print(f"Ошибка добавления примечания методом {method}: {e}")
        return False

def annotate_docx_with_issues(docx_path, pdf_path, analysis_results, output_path):
    """
    ИСПРАВЛЕННАЯ аннотация DOCX файла с поддержкой формул
    
    Args:
        docx_path: путь к исходному DOCX файлу
        pdf_path: путь к PDF файлу (для сопоставления)
        analysis_results: результаты анализа
        output_path: путь для сохранения аннотированного DOCX
    """
    try:
        # Копируем исходный документ
        doc = Document(docx_path)
        
        # Находим изображения в DOCX с их позициями
        docx_images = find_docx_images_with_positions(docx_path)
        
        # НОВОЕ: Находим формулы в DOCX с их позициями
        docx_formulas = find_docx_formulas_with_positions(docx_path)
        
        # Добавляем дополнительные проверки DOCX
        from matcher.docx_checks import check_docx_hyphenation, check_docx_double_spaces, check_docx_margins
        
        docx_issues = []
        docx_issues.extend(check_docx_hyphenation(docx_path))
        docx_issues.extend(check_docx_double_spaces(docx_path))
        docx_issues.extend(check_docx_margins(docx_path))
        
        analysis_results['docx_checks'] = docx_issues
        
        # ИСПРАВЛЕННАЯ логика добавления примечаний к изображениям
        if analysis_results.get("images") and docx_images:
            pdf_images = analysis_results["images"]
            
            # Создаем правильное сопоставление
            image_matching = match_pdf_images_to_docx(pdf_images, docx_images)
            
            print(f"=== ДОБАВЛЕНИЕ ПРИМЕЧАНИЙ К ИЗОБРАЖЕНИЯМ ===")
            
            # Добавляем примечания только к соответствующим изображениям
            for pdf_idx, pdf_img in enumerate(pdf_images):
                if not pdf_img.get('gost_compliant', True):
                    issues = []
                    if not pdf_img.get('is_centered', True):
                        issues.append("изображение не центрировано")
                    if not pdf_img.get('margins_ok', True):
                        issues.append("нарушены поля")
                    if not pdf_img.get('has_empty_line', True):
                        issues.append("отсутствует пустая строка перед изображением")
                    
                    if issues:
                        comment_text = f"📷 ИЗОБРАЖЕНИЕ (стр. {pdf_img['page']}): {'; '.join(issues)}"
                        
                        # Находим соответствующее изображение в DOCX
                        if pdf_idx in image_matching:
                            docx_idx = image_matching[pdf_idx]
                            if docx_idx < len(docx_images):
                                docx_img = docx_images[docx_idx]
                                para_idx = docx_img['paragraph_idx']
                                
                                print(f"Добавляем примечание к изображению {docx_idx+1} в параграфе {para_idx}")
                                try_commenting(doc, para_idx, comment_text, "first")
                            else:
                                print(f"ОШИБКА: docx_idx {docx_idx} вне диапазона")
                        else:
                            print(f"PDF изображение {pdf_idx} не имеет соответствия в DOCX")
        
        # НОВОЕ: Добавляем примечания к формулам
        if analysis_results.get("formulas") and docx_formulas:
            pdf_formulas = analysis_results["formulas"]
            
            # Создаем сопоставление формул
            formula_matching = match_pdf_formulas_to_docx(pdf_formulas, docx_formulas)
            
            print(f"=== ДОБАВЛЕНИЕ ПРИМЕЧАНИЙ К ФОРМУЛАМ ===")
            
            # Добавляем примечания к формулам с нарушениями
            for pdf_idx, pdf_formula in enumerate(pdf_formulas):
                if not pdf_formula.get('gost_compliant', True):
                    issues = []
                    if not pdf_formula.get('is_centered', True):
                        issues.append("формула не центрирована")
                    if not pdf_formula.get('margins_ok', True):
                        issues.append("нарушены поля")
                    if not pdf_formula.get('has_numbering', True):
                        issues.append("отсутствует нумерация")
                    
                    if issues:
                        comment_text = f"🧮 ФОРМУЛА (стр. {pdf_formula['page']}): {'; '.join(issues)} | Текст: '{pdf_formula.get('text', '')[:30]}...'"
                        
                        # Находим соответствующую формулу в DOCX
                        if pdf_idx in formula_matching:
                            docx_idx = formula_matching[pdf_idx]
                            if docx_idx < len(docx_formulas):
                                docx_formula = docx_formulas[docx_idx]
                                para_idx = docx_formula['paragraph_idx']
                                
                                print(f"Добавляем примечание к формуле {docx_idx+1} в параграфе {para_idx}")
                                try_commenting(doc, para_idx, comment_text, "first")
                            else:
                                print(f"ОШИБКА: docx_idx {docx_idx} вне диапазона")
                        else:
                            print(f"PDF формула {pdf_idx} не имеет соответствия в DOCX")
        
        # Добавляем примечания к таблицам (БЕЗ ИЗМЕНЕНИЙ)
        if analysis_results.get("tables"):
            tables_data = analysis_results["tables"]
            
            for table in tables_data:
                if not table.get('gost_compliant', True):
                    issues = []
                    if not table.get('is_centered', True):
                        issues.append("таблица не центрирована")
                    if not table.get('margins_ok', True):
                        issues.append("нарушены поля")
                    if not table.get('has_title', True):
                        issues.append("отсутствует заголовок")
                    if not table.get('has_numbering', True):
                        issues.append("отсутствует нумерация")
                    
                    if issues:
                        comment_text = f"📋 ТАБЛИЦА (стр. {table['page']}): {'; '.join(issues)}"
                        
                        # Ищем параграф перед таблицей
                        table_idx = table.get('table_num', 1) - 1
                        if table_idx < len(doc.tables):
                            # Находим параграф, который идет перед таблицей
                            for paragraph_idx, paragraph in enumerate(doc.paragraphs):
                                if "таблица" in paragraph.text.lower():
                                    try_commenting(doc, paragraph_idx, comment_text, "last")
                                    break
        
        # Остальные примечания (БЕЗ ИЗМЕНЕНИЙ)
        if analysis_results.get("text"):
            text_issues = analysis_results["text"]
            
            # Группируем проблемы по страницам
            page_issues = {}
            for issue in text_issues:
                page = issue.get('page', 1)
                if page not in page_issues:
                    page_issues[page] = []
                page_issues[page].append(issue)
            
            # Добавляем примечания для каждой страницы с проблемами
            for page, issues in page_issues.items():
                if len(issues) > 3:  # Только если много проблем на странице
                    font_issues = sum(1 for i in issues if i.get('issue') == 'font_size')
                    margin_issues = sum(1 for i in issues if i.get('issue') == 'margins')
                    
                    problems = []
                    if font_issues > 0:
                        problems.append(f"размер шрифта ({font_issues} случаев)")
                    if margin_issues > 0:
                        problems.append(f"поля ({margin_issues} случаев)")
                    
                    if problems:
                        comment_text = f"📄 ТЕКСТ (стр. {page}): проблемы с {'; '.join(problems)}"
                        
                        # Добавляем к первому параграфу страницы (примерно)
                        para_idx = max(0, (page - 1) * 20)
                        if para_idx < len(doc.paragraphs):
                            try_commenting(doc, para_idx, comment_text, "first")
        
        # Добавляем примечания по нумерации страниц
        if analysis_results.get("page_numbering"):
            numbering_issues = analysis_results["page_numbering"]
            
            non_compliant_pages = [p for p in numbering_issues if not p.get('gost_compliant', True)]
            if len(non_compliant_pages) > 2:  # Если много проблем с нумерацией
                comment_text = f"📄 НУМЕРАЦИЯ СТРАНИЦ: проблемы на {len(non_compliant_pages)} страницах"
                
                # Добавляем к первому параграфу документа
                if len(doc.paragraphs) > 0:
                    try_commenting(doc, 0, comment_text, "first")
        
        # Добавляем примечания по проверкам DOCX
        for issue in docx_issues:
            issue_type = issue.get('type', '')
            description = issue.get('description', '')
            para_index = issue.get('para_index')
            
            if issue_type == 'hyphenation':
                comment_text = f"🔤 ПЕРЕНОСЫ: {description}"
            elif issue_type == 'soft_hyphen':
                comment_text = f"🔤 МЯГКИЕ ПЕРЕНОСЫ: {description}"
            elif issue_type == 'double_spaces':
                comment_text = f"⎵ ПРОБЕЛЫ: {description}"
            elif 'margin' in issue_type:
                comment_text = f"📏 ПОЛЯ: {description}"
            else:
                comment_text = f"⚠️ DOCX: {description}"
            
            # Добавляем примечание к конкретному параграфу
            if para_index is not None and para_index < len(doc.paragraphs):
                try_commenting(doc, para_index, comment_text, "first")
            else:
                # Добавляем к первому параграфу
                if len(doc.paragraphs) > 0:
                    try_commenting(doc, 0, comment_text, "first")
        
        # Добавляем общую сводку в начало документа
        total_issues = (
            len([img for img in analysis_results.get('images', []) if not img.get('gost_compliant', True)]) +
            len([table for table in analysis_results.get('tables', []) if not table.get('gost_compliant', True)]) +
            len([formula for formula in analysis_results.get('formulas', []) if not formula.get('gost_compliant', True)]) +
            len(analysis_results.get('text', [])) +
            len([page for page in analysis_results.get('page_numbering', []) if not page.get('gost_compliant', True)]) +
            len(analysis_results.get('docx_checks', []))
        )
        
        if total_issues > 0:
            summary_text = f"📊 СВОДКА АНАЛИЗА ГОСТ: Найдено {total_issues} нарушений. Изображения: {len([img for img in analysis_results.get('images', []) if not img.get('gost_compliant', True)])}, Таблицы: {len([table for table in analysis_results.get('tables', []) if not table.get('gost_compliant', True)])}, Формулы: {len([formula for formula in analysis_results.get('formulas', []) if not formula.get('gost_compliant', True)])}, Текст: {len(analysis_results.get('text', []))}, DOCX: {len(analysis_results.get('docx_checks', []))}"
            
            if len(doc.paragraphs) > 0:
                try_commenting(doc, 0, summary_text, "first")
        
        # Сохраняем аннотированный документ
        doc.save(output_path)
        return True
        
    except Exception as e:
        print(f"Ошибка при аннотировании DOCX: {e}")
        return False
