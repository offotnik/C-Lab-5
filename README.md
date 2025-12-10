# C-Lab-5
Лабораторная работа по C# №5 вариант 10

## Задание 

<img width="782" height="456" alt="image" src="https://github.com/user-attachments/assets/817001fb-4da6-4df8-8801-fa207a93b2a2" />

Программа представляет собой консольное приложение для управления данными о движении товаров в сети магазинов. 
Система использует Aspose.Cells для работы с Excel файлами и поддерживает основные CRUD-операции с данными.

### Основные методы

ReadFromExcel() - открывает Excel-файл, читает строки и создает объекты ProductMovement

1. Очистить список movements
2. Загрузить Workbook из filePath
3. Взять первый Worksheet
4. Для каждой строки (начиная со 2-й):
   5. Если первая ячейка пуста → выйти
   6. Считать 6 значений из ячеек
   7. Создать ProductMovement из значений
   8. Добавить в список movements
9. Вывести количество загруженных записей

SaveToExcel() - создает новый Excel-файл, заполняет данными, применяет стили, сохраняет

1. Создать новый Workbook
2. Взять первый Worksheet, назвать "Движение товаров"
3. Записать заголовки в первую строку
4. Применить стиль к заголовкам
5. Для каждого movement в списке:
   6. Записать 6 значений в следующую строку
   7. Применить формат даты к ячейке с датой
8. Автоподобрать ширину колонок
9. Сохранить Workbook в filePath

GetMovementsByStore() - фильтрует операции по ID магазина → список

1. Отфильтровать movements по StoreId == storeId
2. Отсортировать по Date, затем по Article
3. Вывести количество найденных записей
4. Вывести все записи

GetTotalSalesRevenue() - считает сумму: кол-во × цена для всех продаж → одно число

1. Соединить movements и products по Article (join)
2. Отфильтровать только "Продажа"
3. Для каждой записи: Quantity × Price
4. Просуммировать все произведения
5. Вернуть сумму

GetStoreSalesSummary() - группирует продажи по магазинам: сколько и на какую сумму продал каждый → таблица

1. Соединить movements, stores и products (двойной join)
2. Отфильтровать только "Продажа"
3. Сгруппировать по StoreId и StoreName
4. Для каждой группы:
   5. TotalSales = сумма Quantity
   6. TotalRevenue = сумма (Quantity × Price)
7. Отсортировать по TotalRevenue (убывание)
8. Вывести результат

GetMostProfitableProduct() - находит товар с максимальной выручкой от продаж → один товар

1. Соединить movements, products и stores (двойной join)
2. Отфильтровать только "Продажа"
3. Сгруппировать по Article и ProductName
4. Для каждой группы:
   5. TotalRevenue = сумма (Quantity × Price)
   6. TotalSold = сумма Quantity
5. Отсортировать по TotalRevenue (убывание)
6. Взять первую запись → самый прибыльный товар

DeleteMovement() - ищет операцию по ID и удаляет из списка

1. Найти первый movement с OperationId == operationId
2. Если найден:
   3. Удалить из списка movements
   4. Вернуть true
3. Иначе:
   4. Сообщить "не найдено"
   5. Вернуть false

AddMovement() - добавляет новую операцию, проверяя уникальность ID

1. Проверить, есть ли уже movement с таким же OperationId
2. Если есть:
   3. Сообщить "уже существует"
   4. Выйти
3. Иначе:
   4. Добавить movement в список
   5. Сообщить "успешно добавлено"


### Интерфейс

<img width="302" height="181" alt="image" src="https://github.com/user-attachments/assets/50a24142-47b4-44d0-9ea3-37835e9db416" />

### Классы

<img width="1170" height="588" alt="image" src="https://github.com/user-attachments/assets/069152cd-eb64-4117-ac5c-880d329bc403" />

<img width="750" height="455" alt="image" src="https://github.com/user-attachments/assets/53abecbb-0b6a-4058-ba90-903691ddae3f" />

<img width="893" height="498" alt="image" src="https://github.com/user-attachments/assets/39138658-34c7-4707-94de-aacdb3722e55" />

### Тест загрузки данных из XLS

<img width="451" height="231" alt="image" src="https://github.com/user-attachments/assets/edd5a863-8c90-420f-b5c5-e993b7727247" />

### Тест просмтора данных из файла

<img width="959" height="478" alt="image" src="https://github.com/user-attachments/assets/94c5b865-6186-4cdc-86c1-256c81695f8d" />

### Удаление записи

<img width="345" height="251" alt="image" src="https://github.com/user-attachments/assets/fc343dee-2976-43e9-87cf-cb88487145dc" />

#### было 

<img width="500" height="278" alt="image" src="https://github.com/user-attachments/assets/22cd3c00-95b7-47eb-80fc-31a3be4a0aa6" />

#### стало

<img width="619" height="271" alt="image" src="https://github.com/user-attachments/assets/adcfe3bc-e65a-48f1-a8f5-5f932739dff2" />

### Добавление 

<img width="430" height="348" alt="image" src="https://github.com/user-attachments/assets/4fb60541-f287-44e7-9002-8d1d0e218b60" />

#### было

<img width="539" height="487" alt="image" src="https://github.com/user-attachments/assets/6fdbb201-481f-4206-8cc0-bc5e2da7f615" />

#### стало

<img width="668" height="325" alt="image" src="https://github.com/user-attachments/assets/0d2e32e6-4c5f-45e9-9d54-4610a2d2f6f5" />
