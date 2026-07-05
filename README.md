# Financial Dashboard Generator

Автономный генератор финансового HTML-отчета по двум Excel-файлам ОПУ.

## Структура

```text
project/
├── generate_report.py
├── requirements.txt
├── README.md
├── data/
│   ├── m1.xlsx
│   └── m3.xlsx
└── output/
    └── financial_dashboard.html
```

## Запуск

```bash
python -m pip install -r requirements.txt
python generate_report.py
```

После выполнения откройте файл:

```text
output/financial_dashboard.html
```

HTML самодостаточный: стили, данные, JavaScript и Plotly встраиваются внутрь файла.

## Что делает генератор

- автоматически находит месяцы в Excel;
- читает финансовые показатели и статьи затрат;
- строит вкладки `M1`, `M3`, `M1 + M3`, `WHAT IF`;
- формирует KPI, графики Plotly, аналитическую записку и рекомендации;
- продолжает работу при отсутствии отдельных статей и выводит предупреждения в отчете.
