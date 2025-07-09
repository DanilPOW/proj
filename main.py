import gradio as gr
from analyzer import analyze_pdf, analyze_docx_additional
from annotate import annotate_docx_with_issues, remove_existing_comments
from docx2pdf import convert
import os
import tempfile
import shutil

def gradio_analyze(docx_file):
    if docx_file is None:
        return "Пожалуйста, загрузите DOCX файл", None, None
    
    try:
        print(f"=== НАЧАЛО АНАЛИЗА ФАЙЛА: {docx_file.name} ===")
        
        # КРИТИЧЕСКИ ВАЖНО: Удаляем существующие примечания перед конвертацией
        cleaned_docx_path = docx_file.name.replace('.docx', '_cleaned.docx')
        print("=== УДАЛЕНИЕ ПРИМЕЧАНИЙ ===")
        remove_success = remove_existing_comments(docx_file.name, cleaned_docx_path)
        if remove_success:
            print("✓ Примечания успешно удалены")
        else:
            print("⚠️ Проблемы при удалении примечаний, используем исходный файл")
            # Если удаление не удалось, используем исходный файл
            shutil.copy2(docx_file.name, cleaned_docx_path)
        
        # Создаем временный файл для PDF
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
            pdf_path = temp_pdf.name
        
        # Создаем постоянный файл для PDF (для скачивания)
        pdf_output_path = docx_file.name.replace('.docx', '_converted.pdf')
        
        # Конвертируем очищенный DOCX в PDF
        print("=== КОНВЕРТАЦИЯ В PDF ===")
        print(f"Конвертируем: {cleaned_docx_path} -> {pdf_path}")
        
        try:
            convert(cleaned_docx_path, pdf_path)
            # Копируем PDF для скачивания
            shutil.copy2(pdf_path, pdf_output_path)
            print("✓ Конвертация завершена успешно")
        except Exception as conv_error:
            print(f"ОШИБКА КОНВЕРТАЦИИ: {conv_error}")
            
            # Если конвертация не удалась, пробуем с исходным файлом
            print("Пробуем конвертировать исходный файл...")
            try:
                convert(docx_file.name, pdf_path)
                shutil.copy2(pdf_path, pdf_output_path)
                print("✓ Конвертация исходного файла успешна")
            except Exception as conv_error2:
                return f"Ошибка конвертации DOCX в PDF: {conv_error2}", None, None
        
        # Анализируем PDF
        print("=== НАЧАЛО АНАЛИЗА PDF ===")
        results = analyze_pdf(pdf_path)
        print("=== АНАЛИЗ PDF ЗАВЕРШЕН ===")
        
        # Дополнительные проверки DOCX (используем исходный файл для проверок)
        print("=== НАЧАЛО ДОПОЛНИТЕЛЬНЫХ ПРОВЕРОК DOCX ===")
        docx_issues = analyze_docx_additional(docx_file.name)  # Используем исходный файл
        results['docx_checks'] = docx_issues
        print("=== ДОПОЛНИТЕЛЬНЫЕ ПРОВЕРКИ DOCX ЗАВЕРШЕНЫ ===")
        
        # Создаем аннотированный DOCX файл с примечаниями (используем исходный файл)
        output_docx_path = docx_file.name.replace('.docx', '_annotated.docx')
        print("=== СОЗДАНИЕ АННОТИРОВАННОГО ДОКУМЕНТА ===")
        annotate_success = annotate_docx_with_issues(docx_file.name, pdf_path, results, output_docx_path)
        
        # Удаляем только временный PDF, оставляем копию для скачивания
        os.unlink(pdf_path)
        os.unlink(cleaned_docx_path)
        
        # Форматируем результаты
        report = []
        
        # Анализ изображений
        if results.get("images"):
            images_data = results["images"]
            total_images = len(images_data)
            gost_compliant = sum(1 for img in images_data if img.get('gost_compliant', False))
            
            report.append("=== АНАЛИЗ ИЗОБРАЖЕНИЙ ===")
            report.append(f"Всего изображений: {total_images}")
            report.append(f"Соответствуют ГОСТ: {gost_compliant}")
            report.append(f"Не соответствуют ГОСТ: {total_images - gost_compliant}")
            report.append("")
            
            # Детали по каждому изображению
            for i, img in enumerate(images_data[:15], 1):  # Показываем первые 15
                report.append(f"Изображение {i} (страница {img['page']}, тип: {img['type']}):")
                report.append(f"  - ГОСТ: {'✓' if img.get('gost_compliant') else '✗'}")
                report.append(f"  - Центрировано: {'✓' if img.get('is_centered') else '✗'}")
                report.append(f"  - Отступы: {'✓' if img.get('margins_ok') else '✗'}")
                report.append(f"  - Пустая строка: {'✓' if img.get('has_empty_line') else '✗'}")
                report.append(f"  - Размер: {img.get('width_cm', 0):.1f}x{img.get('height_cm', 0):.1f} см")
                if img.get('xref'):
                    report.append(f"  - XRef: {img['xref']}")
                report.append("")
        
        # Анализ таблиц
        if results.get("tables"):
            tables_data = results["tables"]
            total_tables = len(tables_data)
            gost_compliant = sum(1 for table in tables_data if table.get('gost_compliant', False))
            
            report.append("=== АНАЛИЗ ТАБЛИЦ ===")
            report.append(f"Всего таблиц: {total_tables}")
            report.append(f"Соответствуют ГОСТ: {gost_compliant}")
            report.append(f"Не соответствуют ГОСТ: {total_tables - gost_compliant}")
            report.append("")
        
        # Анализ формул
        if results.get("formulas"):
            formulas_data = results["formulas"]
            total_formulas = len(formulas_data)
            gost_compliant = sum(1 for formula in formulas_data if formula.get('gost_compliant', False))
            
            report.append("=== АНАЛИЗ ФОРМУЛ ===")
            report.append(f"Всего формул: {total_formulas}")
            report.append(f"Соответствуют ГОСТ: {gost_compliant}")
            report.append(f"Не соответствуют ГОСТ: {total_formulas - gost_compliant}")
            report.append("")
        
        # Анализ текста
        if results.get("text"):
            text_data = results["text"]
            total_issues = len(text_data)
            
            report.append("=== АНАЛИЗ ТЕКСТА ===")
            report.append(f"Найдено проблем с текстом: {total_issues}")
            
            if total_issues > 0:
                font_issues = sum(1 for issue in text_data if issue.get('issue') == 'font_size')
                margin_issues = sum(1 for issue in text_data if issue.get('issue') == 'margins')
                font_family_issues = sum(1 for issue in text_data if issue.get('issue') == 'font_family')
                
                report.append(f"  - Проблемы с размером шрифта: {font_issues}")
                report.append(f"  - Проблемы с полями: {margin_issues}")
                report.append(f"  - Проблемы с типом шрифта: {font_family_issues}")
            report.append("")
        
        # Анализ нумерации страниц
        if results.get("page_numbering"):
            numbering_data = results["page_numbering"]
            total_pages = len(numbering_data)
            gost_compliant = sum(1 for page in numbering_data if page.get('gost_compliant', False))
            
            report.append("=== АНАЛИЗ НУМЕРАЦИИ СТРАНИЦ ===")
            report.append(f"Всего страниц: {total_pages}")
            report.append(f"Соответствуют ГОСТ: {gost_compliant}")
            report.append(f"Не соответствуют ГОСТ: {total_pages - gost_compliant}")
            report.append("")
        
        # Проверки DOCX
        if results.get("docx_checks"):
            docx_issues = results["docx_checks"]
            total_docx_issues = len(docx_issues)
            
            report.append("=== ДОПОЛНИТЕЛЬНЫЕ ПРОВЕРКИ DOCX ===")
            report.append(f"Найдено проблем: {total_docx_issues}")
            
            if total_docx_issues > 0:
                hyphen_issues = sum(1 for issue in docx_issues if issue.get('type') in ['hyphenation', 'soft_hyphen'])
                space_issues = sum(1 for issue in docx_issues if issue.get('type') == 'double_spaces')
                margin_issues = sum(1 for issue in docx_issues if 'margin' in issue.get('type', ''))
                
                report.append(f"  - Проблемы с переносами: {hyphen_issues}")
                report.append(f"  - Проблемы с пробелами: {space_issues}")
                report.append(f"  - Проблемы с полями: {margin_issues}")
            report.append("")
        
        # Общая сводка
        total_issues = (
            len([img for img in results.get('images', []) if not img.get('gost_compliant', True)]) +
            len([table for table in results.get('tables', []) if not table.get('gost_compliant', True)]) +
            len([formula for formula in results.get('formulas', []) if not formula.get('gost_compliant', True)]) +
            len(results.get('text', [])) +
            len([page for page in results.get('page_numbering', []) if not page.get('gost_compliant', True)]) +
            len(results.get('docx_checks', []))
        )
        
        report.append("=== ОБЩАЯ СВОДКА ===")
        report.append(f"Всего найдено нарушений ГОСТ: {total_issues}")
        report.append(f"Изображений найдено: {len(results.get('images', []))}")
        report.append(f"Таблиц найдено: {len(results.get('tables', []))}")
        report.append(f"Формул найдено: {len(results.get('formulas', []))}")
        
        # Добавляем информацию о PDF в отчет
        report.append("=== ФАЙЛЫ ДЛЯ ДИАГНОСТИКИ ===")
        report.append("PDF файл доступен для скачивания для проверки конвертации")
        if remove_success:
            report.append("✓ Примечания были успешно удалены перед конвертацией")
        else:
            report.append("⚠️ Примечания могли остаться в PDF (проверьте PDF файл)")
        report.append("")
        
        report_text = "\n".join(report) if report else "Анализ завершен, но результаты не найдены."
        
        print("=== АНАЛИЗ ПОЛНОСТЬЮ ЗАВЕРШЕН ===")
        
        if annotate_success and os.path.exists(output_docx_path):
            return report_text, output_docx_path, pdf_output_path
        else:
            return report_text + "\n\nОшибка при создании аннотированного файла.", None, pdf_output_path if os.path.exists(pdf_output_path) else None
        
    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        return f"Ошибка при анализе документа: {str(e)}", None, None

# Создаем интерфейс Gradio
iface = gr.Interface(
    fn=gradio_analyze,
    inputs=gr.File(label="Загрузите DOCX файл", file_types=[".docx"]),
    outputs=[
        gr.Textbox(label="Результат анализа", lines=50, max_lines=100),
        gr.File(label="Скачать аннотированный DOCX с примечаниями"),
        gr.File(label="Скачать конвертированный PDF (проверка конвертации)")
    ],
    title="Анализатор DOCX на соответствие ГОСТ (ИСПРАВЛЕНА ПРОБЛЕМА С ПОВРЕЖДЕНИЕМ ФАЙЛОВ)",
    description="Загрузите DOCX файл для полного анализа на соответствие требованиям ГОСТ. ИСПРАВЛЕНО: Более безопасное удаление примечаний, которое не повреждает DOCX файлы."
)

if __name__ == "__main__":
    iface.launch(share=True)
