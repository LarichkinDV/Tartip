# REPLICATION_GUIDE.md
**Как воспроизвести логику компонента без доступа к исходному коду**  
Версия: 1.0  
Автор: Дмитрий <Фамилия>  
Дата: <дата>

> Документ описывает практический способ повторить поведение компонента (фильтрация → сортировка → группировка → «Итого») **без исходников макроса**.  
> Даны два пути: **A) Excel Power Query (GUI-подход)** и **B) Технологически-нейтральный алгоритм** (псевдокод/ETL), пригодный для реализации на Python/.NET/SQL или в собственном VBA.

---

## 0. Цель
Из исходной табличной выгрузки ИМОКС получить ведомость с промежуточными итогами по группам:
- отфильтровать строки по допустимым парам стадий и непустому типу;
- нормализовать и показать только нужные столбцы (+ все `new_*`);
- отсортировать по Stage → Type → Bucket;
- сформировать строки «Итого» над данными группы;
- **суммировать** только выбранные `*: Double`, **уникально склеивать** прочие поля (в т.ч. несуммируемые `*: Double`, при одном токене — писать числом);
- округление до 7 знаков, «прилипание к целому» (EPS=5e-7), корректные форматы.

См. детализацию требований: `TZ.md`, `DATA_SCHEMAS.md`, `ARCHITECTURE.md`.

---

## 1. Входные данные и требования к схеме

Минимально необходимые столбцы:
- `Type Name : String` (обязательно, непустой),
- `Phase Demolished : String` (обязательно),
- `Phase Created : String` (обязательно).

Опционально:
- `Area : Double` — включает разбиение по диапазонам площади **только** для типа  
  `(потолок)_жилье_натяжной.отм.3м_толщ=5мм`.

Служебный:
- `New_Count : Double` — если отсутствует, его нужно **добавить** и заполнить `1` для всех строк-данных (не «Итого»).

Показывать (whitelist), если присутствуют:
ID; Type Name : String; Category : String; New_Count : Double; Volume : Double;
Area : Double; Length : Double; Width : Double; Phase Demolished : String;
Phase Created : String; Thickness : Double; Perimeter : Double;
Unconnected Height : Double; Height : Double
И **все** столбцы с префиксом `new_` — всегда видимы и **не перезаписываются** логикой итогов (кроме служебного `New_Count : Double`).

---

## 2. Путь A — Excel Power Query (рекомендуемый без кода)

### 2.1 Загрузка и нормализация
1. **Данные → Получить данные → Из таблицы/диапазона** (выделите заголовки на строке 1).  
2. В Power Query:
   - Для всех заголовков примените **Trim**, **Clean**, приведите к нижнему регистру (Transform → Format: *Trim*, *Clean*, *Lower*).  
     > Это имитирует `NormHeader` из `DATA_SCHEMAS.md` (см. §2).
   - При необходимости переименуйте в канон (например, `type name : string`, `phase demolished : string`, и т.д.).
   - Добавьте столбец `new_count : double` (Add Column → Custom) со значением `1`.  
     Преобразуйте его тип в *Decimal Number*.

### 2.2 Фильтрация строк
- Удалите строки, где `type name : string` пуст (Filter → Remove Empty).  
- Оставьте только пары стадий:
  - `phase demolished : string` = `"демонтаж"` **и** `phase created : string` = `"существующие"`;
  - `phase demolished : string` = `"none"` **и** `phase created : string` = `"новая конструкция"`.

> Строки «Итого: …» в исходнике (если есть) нужно исключить из данных (фильтр по `NOT Text.StartsWith([type name : string], "Итого:")`).

### 2.3 Корзины площади (только спец-тип)
Добавьте столбец **Bucket** по условию:
- Если `type name : string` = `(потолок)_жилье_натяжной.отм.3м_толщ=5мм` **и** `area : double` задано, то:
  - `1` если `area < 10 − EPS`,
  - `2` если `area > 10 + EPS AND area < 50 − EPS`,
  - `3` если `area > 50 + EPS`,
  - иначе `0`.
- Для других типов поставьте `0`.

> В M-редакторе удобно объявить `EPS = 0.0000005` и использовать его в выражениях.  
> *Строгость границ* гарантирует, что ровно `10` и `50` (и почти равные) не попадают в 1/2/3.

### 2.4 Ключи сортировки
Добавьте:
- `StageCode`: 1 для `(демонтаж, существующие)`, 2 для `(none, новая конструкция)`, иначе 99.
- `NameKey`: `Text.Lower(Text.Trim([type name : string]))`.
- `Flag`: 1 (данные).  
Отсортируйте по `(StageCode ↑, NameKey ↑, Bucket ↑)`.

### 2.5 Группировка и агрегаты
**Home → Group By** (Advanced):
- Group by: `type name : string`, `StageCode`, `Bucket`.
- Aggregations:
  - **Сумма** для: `new_count : double`, `volume : double`, `area : double`, `length : double`, `perimeter : double`, `unconnected height : double`.  
  - **Индивидуальная агрегация (текст)** для остальных полей: требуется собрать **уникальные значения** и склеить через `;`.

Варианты:
- *Без кода*: сделайте несколько Group By — одну «суммирующую», вторую — «текстовые» поля (как списки), затем на шаге преобразования списков примените **Remove Duplicates**, **Text.Combine** с `";"`.
- *С кратким M-шагом*: внутри агрегации используйте `Text.Combine(List.Distinct(List.Transform(…)), ";")` с предварительным приведением числовых токенов к тексту согласно локальному разделителю.

### 2.6 «Число или текст» для несуммируемых `*: Double`
Для текстовых агрегатов `*: Double` (напр. `thickness : double`, `height : double`, `width : double`):
- Добавьте вычисляемые столбцы-постобработку:
  - Если строка **не содержит `;`** и **распознаётся числом** → преобразуйте в *Decimal Number*, примените **округление до 7** и «прилипание» (через `Number.Round` до 7, затем при `Number.Abs(x - Number.Round(x, 0)) ≤ EPS` → `Number.Round(x,0)`).
  - Иначе оставьте как текст.

### 2.7 Заголовок «Итого»
Создайте таблицу «итогов» из результатов Group By:
- Столбец `Итого: <Type> [<Stage>] + BucketLabel`:
  - Stage: `[Демонтаж]` для `StageCode=1`, `[Новая конструкция]` для `StageCode=2`;
  - BucketLabel: `(до 10м2)`, `(от 10 до 50м2)`, `(от 50м2)` для 1/2/3 и пусто для `0`.

### 2.8 Сборка итогового листа
В Power Query у вас будут **две таблицы**:
1) **Итоги по группам** (одна строка на группу).  
2) **Данные** (отфильтрованные и отсортированные).

Экспортируйте обе на разные листы.  
На основном листе Excel можно собрать единый вид:
- Вставьте «Итого» **над** строками своей группы: используйте `ВПР/XLOOKUP` или Power Query Merge и затем в Excel разложите по ключам `(StageCode, NameKey, Bucket)`:
  - Сортируйте обе таблицы одинаково.
  - Выполните **объединение**: для каждой группы сначала строка с «Итого», затем — её данные.
- Для визуальной группировки примените **Данные → Структура → Группировать** (ручная группировка под «Итого»).

> Примечание: Power Query не управляет Excel-Outline. Группировку (структуру) выполняйте уже в Excel (можно макрозаписью, если допустимо).

---

## 3. Путь B — Технологически нейтральный алгоритм (для Python/.NET/SQL/VBA)

### 3.1 Константы
- `EPS = 5e-7`.
- Набор суммируемых колонок:  
  `{ "New_Count : Double", "Volume : Double", "Area : Double", "Length : Double", "Perimeter : Double", "Unconnected Height : Double" }`.

### 3.2 Псевдокод
```text
INPUT: table rows with headers
NORMALIZE headers -> lower, trim, unify " : " (see DATA_SCHEMAS.md)

ENSURE "new_count : double": if absent, insert and set = 1 for all data rows

FILTER:
  keep rows where type name not empty
  keep rows where (demol, created) in {("демонтаж","существующие"), ("none","новая конструкция")}
  exclude rows where type name starts with "Итого:"

BUCKET(type, area):
  if type == SPECIAL and area is number:
     if area < 10 - EPS -> 1
     else if area > 10 + EPS and area < 50 - EPS -> 2
     else if area > 50 + EPS -> 3
     else -> 0
  else -> 0

STAGECODE(demol, created):
  ("демонтаж","существующие") -> 1
  ("none","новая конструкция") -> 2
  else -> 99

BUILD KEYS for each row:
  stage := STAGECODE(...)
  namekey := lower(trim(type))
  bucket := BUCKET(type, area)

SORT rows by (stage asc, namekey asc, bucket asc)

GROUP BY (type, stage, bucket):
  For each column:
    if column in SUM_SET:
       sum only numeric values; if all empty/zero -> result empty
       else -> round to 7, snap to int if |x - round(x,0)| <= EPS
    else:
       collect distinct tokens (stringify numbers with local decimal, trim trailing zeros)
       joined := join(distinct, ";")
       if header is "* : Double" and joined has single numeric token:
           write NUMBER (round 7, snap-to-int)
       else:
           write TEXT (joined)

WRITE one "Itogo" row per group BEFORE its data:
  caption := "Итого: " + type + " [" + ("Демонтаж" or "Новая конструкция") + "]" + bucket_label
  where bucket_label is "(до 10м2)" | "(от 10 до 50м2)" | "(от 50м2)" for 1/2/3 else ""

PRESENTATION:
  show whitelist columns + all starting with "new_"
  numeric formats: integer -> "0"; fractional -> "0.#######"
  non-sum "*: Double" with multiple tokens -> text
  non-overwrite columns whose header starts with "new_" (except "new_count : double")

IDEMPOTENCY:
  on rerun, detect existing Itogo by caption and update values only
```

### 3.3 Проверки («красные флаги»)

Отсутствие обязательных колонок.
Попадание 10/50 в корзины (не должно).
«Хвосты» у дробных сумм → после форматирования должны исчезать; целые — без десятичной части.
Несуммируемые *: Double: одиночный токен не должен преобразоваться в текст.

## 4. Пример мини-набора данных (CSV)

input.csv
Type Name : String,Phase Demolished : String,Phase Created : String,Area : Double,Thickness : Double,Volume : Double
(потолок)_жилье_натяжной.отм.3м_толщ=5мм,Демонтаж,Существующие,9.9999998,40,1.25
(потолок)_жилье_натяжной.отм.3м_толщ=5мм,Демонтаж,Существующие,49.9,40,0.75
2ПБ13-1,Демонтаж,Существующие,,39.9999999999,2
2ПБ13-1,None,Новая конструкция,8,40,3

Ожидаемые «Итого» (фрагмент):
Итого: (потолок)_жилье_натяжной.отм.3м_толщ=5мм [Демонтаж] (до 10м2)
  Area : Double  -> 10
  Thickness : Double -> 40   (один токен -> число)
  Volume : Double -> 2.0 (1.25 + 0.75 для обеих групп ниже: см. раздельные бакеты)
Итого: (потолок)_жилье_натяжной.отм.3м_толщ=5мм [Демонтаж] (от 10 до 50м2)
  Area : Double  -> 49.9
  Thickness : Double -> 40
Итого: 2ПБ13-1 [Демонтаж]
  Thickness : Double -> 40 (39.9999999999 -> округл. до 40, прилипание)
Итого: 2ПБ13-1 [Новая конструкция]
  Area : Double -> 8
  Thickness : Double -> 40
Внимание: в реальном файле итог Volume будет по каждой группе отдельно; пример выше демонстрирует правила округления/записи.

## 5. Проверочный чек-лист (минимум)

[] После нормализации заголовков обязательные поля распознаются.
[] Пустые Type Name : String удалены.
[] В таблице остались только две пары стадий.
[] Для спец-типа разложение по корзинам: <10, >10 & <50, >50; ровно 10/50 не попали в корзины.
[] Суммы: только по new_count, volume, area, length, perimeter, unconnected height.
[] Несуммируемые *: Double: один токен → число; несколько → текст "v1;v2;…".
[] «Хвосты» нулей и «висящие» разделители отсутствуют; локаль учтена.
[] На повторном запуске (или повторной прогонке ETL) «Итого» обновляются, а new_* (кроме new_count) не перезаписываются.

## 6. Частые вопросы (FAQ)

Q: Что делать, если отсутствует Area : Double?
A: Корзины площади не применяются; все группы идут с Bucket=0 без метки диапазона.

Q: Нужно ли учитывать регистр/пробелы/варианты двоеточия в заголовках?
A: Да — обязательно нормализуйте имена по правилам NormHeader (см. DATA_SCHEMAS.md §2.2).

Q: Как добиться «прилипания» к целому?
A: После округления до 7 знаков, если |x − round(x)| ≤ 5e-7, используйте целое round(x) и формат "0".

Q: Что с колонками new_*?
A: Они всегда видимы и не перезаписываются итогами (кроме служебного New_Count : Double).

Q: Можно ли реализовать в SQL?
A: Да: нормализация заголовков на стадии загрузки, фильтр по стадиям, CASE для корзин, GROUP BY с SUM для нужных колонок, STRING_AGG(DISTINCT …, ';') для остальных и постобработка «один токен → число».

## 7. Что считать успешной репликацией

Поведение полностью соответствует TZ.md / DATA_SCHEMAS.md.
Итоги на тех же входных данных равны по значениям и форматам (с учётом EPS).
Визуальный вид: «Итого» над данными группы; whitelist столбцов и все new_* видны; лишние скрыты/отсутствуют.
Повторный прогон приводит к тем же результатам (идемпотентность).

## 8. Примечания по эксплуатации

Excel/Power Query: группировку (Outline) делайте на финальном листе вручную/полуавтоматом, PQ его не создаёт.
Большие объёмы: для >100k строк рассмотрите вариант B (скрипт/сервис) и выгрузку в CSV/Parquet.
Кодировки: храните документы в UTF-8, чтобы не получить артефакты "Итого: …".

## 9. Связанные документы

TZ.md — формальные требования и алгоритм.
ARCHITECTURE.md — модули и потоки.
DATA_SCHEMAS.md — схема полей и нормализация.
TEST_PLAN.md — эталонные кейсы и метрики.
