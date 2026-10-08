-- Схема КХД для бюджетных форм (И1, И2, И34, И35, КОНС1)
-- Запуск: clickhouse-client --multiquery < schema.sql
--   или:  python load_to_clickhouse.py --create-only

CREATE DATABASE IF NOT EXISTS budget;

-- Основная (сырая/нормализованная) таблица: одна строка = одна статья формы за один период
CREATE TABLE IF NOT EXISTS budget.budget_forms
(
    `id`                UUID DEFAULT generateUUIDv4(),   -- суррогатный ключ строки
    `source_file`       String,                          -- имя исходного Excel-файла
    `sheet_name`        String,                          -- имя листа в файле
    `report_year`       UInt16,                          -- год отчета (из имени файла / "Период отчета")
    `report_type`       LowCardinality(String),          -- План | Факт
    `form`              LowCardinality(String),          -- И1 | И2 | И34 | И35 | КОНС1
    `organization`      String,                          -- организационная единица
    `currency`          LowCardinality(String),          -- основная валюта
    `scenario`          String,                          -- сценарий отчета из шапки
    `row_num`           UInt32,                          -- порядковый номер строки на листе
    `row_code`          String,                          -- "Код статьи" ("№ п/п", если кода нет)
    `row_name`          String,                          -- наименование статьи
    `unit`              String,                          -- единицы измерения
    `ekbs_code`         String,                          -- Статья ЕКБС / Статья АС ФЗД
    `period`            String,                          -- нормализованный период колонки
    `period_kind`       LowCardinality(String),          -- year | quarter | half | nine_month | month | variance | other
    `value`             Float64,                         -- значение статьи за период
    `loaded_at`         DateTime DEFAULT now()
)
ENGINE = MergeTree
PARTITION BY (report_year, form)
ORDER BY (report_year, report_type, form, organization, row_code, period);

-- Материализованное представление "широкой" таблицы: периоды -> колонки Map(String, Float64)
CREATE TABLE IF NOT EXISTS budget.budget_forms_wide
(
    `id`                UUID,
    `source_file`       String,
    `report_year`       UInt16,
    `report_type`       LowCardinality(String),
    `form`              LowCardinality(String),
    `organization`      String,
    `currency`          LowCardinality(String),
    `row_code`          String,
    `row_name`          String,
    `unit`              String,
    `ekbs_code`         String,
    `periods`           Map(String, Float64),
    `loaded_at`         DateTime
)
ENGINE = ReplacingMergeTree(loaded_at)
PARTITION BY (report_year, form)
ORDER BY (report_year, report_type, form, organization, row_code);

INSERT INTO budget.budget_forms_wide
SELECT
    id,
    anyLast(source_file)                        AS source_file,
    report_year,
    report_type,
    form,
    organization,
    currency,
    row_code,
    anyLast(row_name)                           AS row_name,
    anyLast(unit)                               AS unit,
    anyLast(ekbs_code)                          AS ekbs_code,
    mapFromVectors(groupArray(period), groupArray(value)) AS periods,
    max(loaded_at)                              AS loaded_at
FROM budget.budget_forms
GROUP BY id, report_year, report_type, form, organization, currency, row_code;
