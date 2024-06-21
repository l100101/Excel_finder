import pandas as pd

try:
    # Загрузка Excel файла
    file_path = 'SM.xlsx'
    df = pd.read_excel(file_path)
    print("Excel файл успешно загружен.")
    
    # Показать первые несколько строк для проверки
    print("Первые строки DataFrame:")
    print(df.head())
    
    # Запрос названий колонок у пользователя
 
    column_i = input("Введите название колонки, в которой будет производиться поиск знач, содержащих  Com|IP_MK  (например, 'Цепь'): ")
    column_a = input("Введите название колонки, содержащей значения для извлечения  (например, 'PortPin'): ")
       
    # Проверка, что указанные колонки существуют
    if column_i not in df.columns:
        raise ValueError(f"Колонка {column_i} не найдена в файле.")
    if column_a not in df.columns:
        raise ValueError(f"Колонка {column_a} не найдена в файле.")

    # Фильтрация строк, содержащих "Com" или "IP_MK" в указанной колонке
    filtered_df = df[df[column_i].str.contains('Com|IP_MK', na=False, case=False)]
    print(f"Найдено {len(filtered_df)} строк, содержащих 'Com' или 'IP_MK'.")

    # Создание нового DataFrame с указанными колонками
    result_df = filtered_df[[column_a, column_i]]
    print("Новый DataFrame создан.")

    # Сохранение результата в новый Excel файл
    result_file_path = 'filtered_results.xlsx'
    result_df.to_excel(result_file_path, index=False)
    print(f"Результаты сохранены в '{result_file_path}'.")

except Exception as e:
    print(f"Произошла ошибка: {e}")

# Ожидание ввода пользователя перед закрытием
input("Нажмите Enter для завершения.")
