import fitz
import re
from docx import Document
from matcher.images import analyze_pdf_images
from matcher.formulas import analyze_pdf_formulas
from matcher.tables import analyze_pdf_tables
from matcher.text import analyze_pdf_text
from matcher.page_numbering import analyze_page_numbering
from matcher.docx_checks import check_docx_hyphenation, check_docx_double_spaces, check_docx_margins

def analyze_pdf(pdf_path):
    """
    Основная функция полного анализа PDF файла
    ИЗМЕНЕН ПОРЯДОК: сначала формулы, потом изображения
    
    Args:
        pdf_path: путь к PDF файлу
    
    Returns:
        dict: результаты анализа всех элементов
    """
    try:
        results = {}
        
        # СНАЧАЛА анализ формул (БЕЗ исключения областей изображений)
        try:
            print("=== НАЧАЛО АНАЛИЗА ФОРМУЛ (ПРИОРИТЕТНЫЙ) ===")
            formulas_results = analyze_pdf_formulas(pdf_path, [])  # Пустой список изображений
            results["formulas"] = formulas_results
            print(f"=== ЗАВЕРШЕН АНАЛИЗ ФОРМУЛ: {len(formulas_results)} найдено ===")
        except Exception as e:
            print(f"Ошибка при анализе формул: {e}")
            results["formulas"] = []
        
        # ЗАТЕМ анализ изображений (с исключением областей формул)
        try:
            print("=== НАЧАЛО АНАЛИЗА ИЗОБРАЖЕНИЙ ===")
            images_results = analyze_pdf_images(pdf_path)
            results["images"] = images_results
            print(f"=== ЗАВЕРШЕН АНАЛИЗ ИЗОБРАЖЕНИЙ: {len(images_results)} найдено ===")
        except Exception as e:
            print(f"Ошибка при анализе изображений: {e}")
            results["images"] = []
        
        # Анализ таблиц (исключаем области изображений И формул)
        try:
            print("=== НАЧАЛО АНАЛИЗА ТАБЛИЦ ===")
            # Объединяем области изображений и формул для исключения
            excluded_areas = results.get("images", []) + results.get("formulas", [])
            tables_results = analyze_pdf_tables(pdf_path, excluded_areas)
            results["tables"] = tables_results
            print(f"=== ЗАВЕРШЕН АНАЛИЗ ТАБЛИЦ: {len(tables_results)} найдено ===")
        except Exception as e:
            print(f"Ошибка при анализе таблиц: {e}")
            results["tables"] = []
        
        # Анализ текста
        try:
            print("=== НАЧАЛО АНАЛИЗА ТЕКСТА ===")
            text_results = analyze_pdf_text(pdf_path)
            results["text"] = text_results
            print(f"=== ЗАВЕРШЕН АНАЛИЗ ТЕКСТА: {len(text_results)} проблем найдено ===")
        except Exception as e:
            print(f"Ошибка при анализе текста: {e}")
            results["text"] = []
        
        # Анализ нумерации страниц
        try:
            print("=== НАЧАЛО АНАЛИЗА НУМЕРАЦИИ ===")
            numbering_results = analyze_page_numbering(pdf_path)
            results["page_numbering"] = numbering_results
            print(f"=== ЗАВЕРШЕН АНАЛИЗ НУМЕРАЦИИ: {len(numbering_results)} страниц проверено ===")
        except Exception as e:
            print(f"Ошибка при анализе нумерации: {e}")
            results["page_numbering"] = []
        
        return results
        
    except Exception as e:
        print(f"Общая ошибка при анализе PDF: {e}")
        return {
            "images": [],
            "tables": [],
            "formulas": [],
            "text": [],
            "page_numbering": [],
            "docx_checks": []
        }

def analyze_docx_additional(docx_path):
    """
    Дополнительные проверки DOCX файла
    
    Args:
        docx_path: путь к DOCX файлу
    
    Returns:
        list: список проблем
    """
    try:
        issues = []
        
        print("=== НАЧАЛО ДОПОЛНИТЕЛЬНЫХ ПРОВЕРОК DOCX ===")
        
        # Проверка переносов
        hyphen_issues = check_docx_hyphenation(docx_path)
        issues.extend(hyphen_issues)
        print(f"Проверка переносов: {len(hyphen_issues)} проблем")
        
        # Проверка двойных пробелов
        space_issues = check_docx_double_spaces(docx_path)
        issues.extend(space_issues)
        print(f"Проверка пробелов: {len(space_issues)} проблем")
        
        # Проверка полей
        margin_issues = check_docx_margins(docx_path)
        issues.extend(margin_issues)
        print(f"Проверка полей: {len(margin_issues)} проблем")
        
        print(f"=== ЗАВЕРШЕНЫ ДОПОЛНИТЕЛЬНЫЕ ПРОВЕРКИ DOCX: {len(issues)} проблем ===")
        
        return issues
        
    except Exception as e:
        print(f"Ошибка при дополнительных проверках DOCX: {e}")
        return []
