"""Англійські переклади інтерфейсу: ключ – український текст, значення – англійський."""

EN = {
    'λ1, мс⁻¹':
        'λ1, ms⁻¹',
    's1, мс':
        's1, ms',
    's2, мс':
        's2, ms',
    'D1, мс':
        'D1, ms',
    'λ2, мс⁻¹':
        'λ2, ms⁻¹',
    's2b, мс':
        's2b, ms',
    'D2, мс':
        'D2, ms',
    'Сер. очікування ISR':
        'Mean ISR wait',
    'Сер. очікування вторинної, МРЧ':
        'Mean secondary wait, SRT',
    'Сер. очікування вторинної, ЖРЧ':
        'Mean secondary wait, HRT',
    'Сер. час перебування, МРЧ':
        'Mean sojourn time, SRT',
    'Сер. час перебування, ЖРЧ':
        'Mean sojourn time, HRT',
    'Макс. очікування ISR':
        'Max ISR wait',
    'Макс. очікування вторинної, МРЧ':
        'Max secondary wait, SRT',
    'Макс. очікування вторинної, ЖРЧ':
        'Max secondary wait, HRT',
    'Макс. час перебування, МРЧ':
        'Max sojourn time, SRT',
    'Макс. час перебування, ЖРЧ':
        'Max sojourn time, HRT',
    'Сер. довжина черги Q0':
        'Mean queue length Q0',
    'Сер. довжина черги QS (МРЧ)':
        'Mean queue length QS (SRT)',
    'Сер. довжина черги QH (ЖРЧ)':
        'Mean queue length QH (HRT)',
    'Завантаження ISR':
        'ISR utilization',
    'Завантаження вторинної, МРЧ':
        'Secondary utilization, SRT',
    'Завантаження вторинної, ЖРЧ':
        'Secondary utilization, HRT',
    'P порушення дедлайну, МРЧ':
        'P deadline miss, SRT',
    'P порушення дедлайну, ЖРЧ':
        'P deadline miss, HRT',
    'ЖРЧ':
        'HRT',
    'МРЧ':
        'SRT',
    'Сценарій {n}':
        'Scenario {n}',
    'Кількість оброблених подій':
        'Events processed',
    'Сер. очікування первинної обробки (ISR), мс':
        'Mean primary (ISR) wait, ms',
    'Сер. очікування вторинної обробки, мс':
        'Mean secondary wait, ms',
    'Сер. час реакції, мс':
        'Mean response time, ms',
    'Час реакції, 95-й перцентиль, мс':
        'Response time, 95th percentile, ms',
    'Час реакції, 99-й перцентиль, мс':
        'Response time, 99th percentile, ms',
    'Макс. час реакції, мс':
        'Max response time, ms',
    'Ймовірність порушення дедлайну':
        'Deadline miss probability',
    'Завантаження вторинним обробником':
        'Secondary processor utilization',
    'Кількість подій під час відмови':
        'Events during the fault',
    'Сер. час реакції під час відмови, мс':
        'Mean response time during the fault, ms',
    'Макс. час реакції під час відмови, мс':
        'Max response time during the fault, ms',
    'Ймовірність порушення дедлайну під час відмови':
        'Deadline miss probability during the fault',
    '{name}: допустимо від {lo} до {hi}':
        '{name}: allowed from {lo} to {hi}',
    'N має бути цілим числом':
        'N must be an integer',
    'таблиця':
        'table',
    'Програма':
        'Program',
    'Симулятор моделі реактивності {v}':
        'Reactivity model simulator {v}',
    'Дата':
        'Date',
    'Джерело':
        'Source',
    'Кількість прогонів R':
        'Number of replications R',
    'Довірчий інтервал':
        'Confidence interval',
    '95 %, розподіл Стьюдента':
        '95 %, Student distribution',
    'Сценарій':
        'Scenario',
    'U (без відмов)':
        'U (without faults)',
    'Файл у «рядковому» форматі (сценарії – у рядках). Його перетворено: тепер кожен сценарій – окремий стовпець таблиці.':
        'The file is in row format (scenarios in rows). It has been converted: each scenario is now a separate table column.',
    'У файлі немає жодного сценарію з даними.':
        'The file contains no scenario with data.',
    'ПОМИЛКА':
        'ERROR',
    'Попередження':
        'Warning',
    'Параметри':
        'Parameters',
    'Результати':
        'Results',
    'Довірчі_інтервали_95':
        'Confidence_intervals_95',
    'Метадані':
        'Metadata',
    'Тип':
        'Type',
    'Клас':
        'Class',
    'Інтенсивність, 1/мс':
        'Rate, 1/ms',
    'WCET + T_core, мс':
        'WCET + T_core, ms',
    'Дедлайн, мс':
        'Deadline, ms',
    'U під час відмов':
        'U during faults',
    'Відмова':
        'Fault',
    'Типи_подій':
        'Event_types',
    'Файл має неправильний формат: потрібно 8 рядків параметрів і хоча б один стовпець-сценарій (або перший рядок з назвами параметрів lambda1, s1, s2, N, D1, lambda2, s2b, D2).':
        'Wrong file format: 8 parameter rows and at least one scenario column are required (or a first row with parameter names lambda1, s1, s2, N, D1, lambda2, s2b, D2).',
    'не заповнено: {m}':
        'not filled: {m}',
    'Текст':
        'Text',
    'Параметр':
        'Parameter',
    'Значення':
        'Value',
    'Рядок {r}, «{k}»: «{raw}» – не число, комірку очищено.':
        'Row {r}, “{k}”: “{raw}” is not a number, the cell was cleared.',
    'Стовпець {c} (сценарій {n}), рядок {r}: «{raw}» – не число, комірку очищено.':
        'Column {c} (scenario {n}), row {r}: “{raw}” is not a number, the cell was cleared.',
    'Неочікувана помилка':
        'Unexpected error',
    'Повний текст записано у файл {p}':
        'Full text written to {p}',
    '⌂   Початок':
        '⌂   Home',
    '▦   Дані Excel':
        '▦   Excel data',
    '◈   Типи подій і відмови':
        '◈   Event types & faults',
    '≡   Результати':
        '≡   Results',
    '◔   Графіки':
        '◔   Charts',
    '↺   Історія':
        '↺   History',
    '1.  Задайте параметри: відкрийте Excel-файл або сценарій (чи створіть новий).\n2.  Ліворуч унизу задайте кількість прогонів R і seed. Кожен прогін – окрема імітація з seed, seed + 1, …; результат – середнє і 95 % довірчий інтервал.\n3.  Натисніть «Запустити». Розрахунок іде у фоні на всіх ядрах процесора.\n4.  Перегляньте «Результати» і «Графіки», збережіть звіт в Excel або результат у базу даних.\n5.  Повторне збереження тих самих даних не створює дублікатів.':
        '1.  Set the parameters: open an Excel file or a scenario (or create a new one).\n2.  At the bottom left, set the number of replications R and the seed. Each replication is a separate simulation with seed, seed + 1, …; the result is the mean and the 95 % confidence interval.\n3.  Press “Run”. The computation runs in the background on all processor cores.\n4.  View “Results” and “Charts”, save an Excel report or store the result in the database.\n5.  Saving the same data again does not create duplicates.',
    'Як це працює':
        'How it works',
    '  Реактивність БО':
        '  On-board reactivity',
    '  НАЛАШТУВАННЯ ЗАПУСКУ':
        '  RUN SETTINGS',
    'Кількість незалежних прогонів R. Більше – вужчий довірчий інтервал, довше рахує.':
        'Number of independent replications R. More – narrower confidence interval, longer run.',
    'Початкове значення генератора. Той самий seed – ті самі результати.':
        'Random generator seed. The same seed gives the same results.',
    'Змінити мову інтерфейсу (потрібен перезапуск програми)':
        'Change the interface language (the program restarts)',
    'Готово':
        'Ready',
    'Сховище: перевірка…':
        'Storage: checking…',
    'Помилка під час симуляції':
        'Error during simulation',
    'Помилка симуляції':
        'Simulation error',
    'Симуляція серії: {n} сценаріїв × {r} прогонів…':
        'Simulating series: {n} scenarios × {r} replications…',
    'Готово: {n} сценаріїв за {t} с':
        'Done: {n} scenarios in {t} s',
    'Симуляція сценарію «{s}»: {r} прогонів…':
        'Simulating scenario “{s}”: {r} replications…',
    'Готово: сценарій «{s}» за {t} с':
        'Done: scenario “{s}” in {t} s',
    'Симулятор моделі реактивності':
        'Reactivity model simulator',
    'Розрахунок показників реактивності бортового обчислювача: час реакції на переривання, черги, завантаження процесора і ймовірність порушення дедлайнів.':
        'Computing reactivity metrics of an on-board computer: interrupt response time, queues, processor utilization and deadline miss probability.',
    'Серія сценаріїв з Excel':
        'Scenario series from Excel',
    'Таблиця з 8 параметрами в кожному стовпці: два типи подій (МРЧ і ЖРЧ). Підходить для дослідження впливу одного параметра – наприклад, s2 або D1.':
        'A table with 8 parameters in each column: two event types (SRT and HRT). Suitable for studying the effect of one parameter, e.g. s2 or D1.',
    'Відкрити Excel-файл…':
        'Open Excel file…',
    'Сценарій з кількома типами подій':
        'Scenario with several event types',
    'Довільна кількість типів подій, параметри платформи в тактах і сценарії відмов: шторм переривань і перевищення WCET.':
        'Any number of event types, platform parameters in clock cycles and fault scenarios: interrupt storm and WCET overrun.',
    'Відкрити сценарій':
        'Open scenario',
    'Історія результатів':
        'Results history',
    'Збережені в MongoDB результати: пошук, повторне відкриття, порівняння кількох запусків.':
        'Results stored in MongoDB: search, reopen, compare several runs.',
    'Перейти до історії':
        'Go to history',
    'версія {v}':
        'version {v}',
    '☀  Світла тема':
        '☀  Light theme',
    '☾  Темна тема':
        '☾  Dark theme',
    'Симуляція: прогін {a} з {b}  ·  {t} с':
        'Simulation: replication {a} of {b}  ·  {t} s',
    'Прогонів R':
        'Replications R',
    'Сховище результатів: {d}':
        'Results storage: {d}',
    'Сховище: помилка':
        'Storage: error',
    ', залишилось ≈ {t} с':
        ', ≈ {t} s left',
    'Немає жодного сценарію, який можна змоделювати. ':
        'No scenario can be simulated. ',
    'Симуляція ще триває':
        'Simulation is still running',
    'Симуляція ще не завершилась. Закрити програму?':
        'The simulation has not finished. Close the program?',
    'JSON-файли':
        'JSON files',
    ' – MongoDB не запущена, результати зберігаються у файлах.':
        ' – MongoDB is not running, results are stored in files.',
    'Заповніть таблицю або відкрийте Excel-файл.':
        'Fill in the table or open an Excel file.',
    'Результат пораховано, але його не вдалося повністю показати – подробиці у logs/errors.log':
        'The result was computed but could not be fully displayed – see logs/errors.log',
    'Помилки: ':
        'Errors: ',
    'сценарій {n}: {e}':
        'scenario {n}: {e}',
    'Час реакції за типами подій':
        'Response time by event type',
    'Час реакції в часі (один прогін)':
        'Response time over time (one replication)',
    'Ймовірність порушення дедлайнів':
        'Deadline miss probability',
    'Час очікування від тривалості вторинної обробки s2':
        'Waiting time vs. secondary processing time s2',
    'Середній час очікування, мс':
        'Mean waiting time, ms',
    'Максимальний час очікування від s2':
        'Maximum waiting time vs. s2',
    'Максимальний час очікування, мс':
        'Maximum waiting time, ms',
    'Довжина черг від s2':
        'Queue lengths vs. s2',
    'Середня довжина черги':
        'Mean queue length',
    'Час перебування (реакції) від s2':
        'Sojourn (response) time vs. s2',
    'Середній час перебування, мс':
        'Mean sojourn time, ms',
    'Завантаження процесора від s2':
        'Processor utilization vs. s2',
    'Коефіцієнт завантаження':
        'Utilization',
    'Порушення дедлайнів МРЧ від дедлайну D1':
        'SRT deadline misses vs. deadline D1',
    'Порушення дедлайнів ЖРЧ від дедлайну D2':
        'HRT deadline misses vs. deadline D2',
    'Кожна точка – найбільший час реакції подій цього типу за короткий інтервал часу. Суцільні лінії – ЖРЧ, штрихові – МРЧ.':
        'Each point is the largest response time of this event type within a short time interval. Solid lines – HRT, dashed – SRT.',
    'Первинна обробка (ISR)':
        'Primary processing (ISR)',
    'Вторинна обробка, МРЧ':
        'Secondary processing, SRT',
    'Вторинна обробка, ЖРЧ':
        'Secondary processing, HRT',
    'Черга ISR (Q0)':
        'ISR queue (Q0)',
    'Черга МРЧ (QS)':
        'SRT queue (QS)',
    'Черга ЖРЧ (QH)':
        'HRT queue (QH)',
    'Ймовірність порушення дедлайну МРЧ':
        'SRT deadline miss probability',
    'Ймовірність порушення дедлайну ЖРЧ':
        'HRT deadline miss probability',
    'Графіки':
        'Charts',
    'Спочатку запустіть серію або сценарій.':
        'Run a series or a scenario first.',
    'Довірчі інтервали':
        'Confidence intervals',
    'Логарифмічна шкала':
        'Log scale',
    'Точки – середнє за R прогонів; вертикальні риски – 95 % довірчий інтервал.':
        'Points are means over R replications; vertical bars are 95 % confidence intervals.',
    'Час реакції, мс':
        'Response time, ms',
    'Час реакції за типами подій (середнє за R прогонів)':
        'Response time by event type (mean over R replications)',
    'Червона риска – дедлайн типу. Якщо стовпчик «максимальний» заходить за риску, дедлайн порушувався.':
        'The red tick is the deadline of the type. If the “maximum” bar crosses the tick, the deadline was missed.',
    'Модельний час, с':
        'Model time, s',
    'Макс. час реакції в інтервалі, мс':
        'Max response time in interval, ms',
    ' Червона смуга – інтервал дії відмови.':
        ' The red band is the fault interval.',
    'Для ЖРЧ будь-яке значення більше нуля означає порушення жорсткої вимоги.':
        'For HRT any value above zero means a hard requirement was violated.',
    '✥ Перемістити':
        '✥ Pan',
    'Перетягуйте графік мишею (ще раз – вимкнути)':
        'Drag the chart with the mouse (click again to turn off)',
    '⊕ Збільшити':
        '⊕ Zoom',
    'Виділіть мишею прямокутник, щоб його збільшити (ще раз – вимкнути)':
        'Select a rectangle with the mouse to zoom in (click again to turn off)',
    'Зберегти графік':
        'Save chart',
    'Зберегти PNG…':
        'Save PNG…',
    'Показник':
        'Metric',
    'Вісь X':
        'X axis',
    'Серія «{s}»':
        'Series “{s}”',
    'Графік':
        'Chart',
    'Довільний графік: будь-який показник від будь-якого параметра':
        'Custom chart: any metric vs. any parameter',
    '№ сценарію':
        'Scenario no.',
    'немає змодельованих сценаріїв':
        'no simulated scenarios',
    'немає додатних значень для логарифмічної шкали':
        'no positive values for the log scale',
    '✕ – перевантаження (U ≥ 1): черга росте без меж, ці значення залежать від N і не є стаціонарними – для висновків їх не використовують.':
        '✕ – overload (U ≥ 1): the queue grows without bound, these values depend on N and are not steady-state – do not use them for conclusions.',
    'Точки з імовірністю 1 – дедлайн менший за саму обробку (s1 + s2), тобто недосяжний.':
        'Points with probability 1 – the deadline is shorter than the processing itself (s1 + s2), i.e. unreachable.',
    'Лінії «{t}» збігаються, тому показано одну.':
        'Lines “{t}” coincide, so only one is shown.',
    'Лінії {t} майже збігаються – друга показана штрихом.':
        'Lines {t} almost coincide – the second one is dashed.',
    'Тривалість вторинної обробки МРЧ s2, мс':
        'SRT secondary processing time s2, ms',
    '{p} на ці показники не впливає, тому сценарії, що відрізняються лише цим параметром, дають однакові результати і показані однією точкою.':
        '{p} does not affect these metrics, so scenarios differing only in this parameter give identical results and are shown as one point.',
    '{p} на ці показники не впливають, тому сценарії, що відрізняються лише ними, дають однакові результати і показані однією точкою.':
        '{p} do not affect these metrics, so scenarios differing only in them give identical results and are shown as one point.',
    'Не показано сценаріїв: {n} – у них {p} не змінювався (одна точка не утворює лінії).':
        'Scenarios not shown: {n} – {p} does not vary in them (one point does not make a line).',
    'Показник практично не залежить від {p} – лінія горизонтальна, це очікуваний результат.':
        'The metric practically does not depend on {p} – the line is horizontal, this is the expected result.',
    'R = {r}: інтервали дуже широкі, для звіту беріть R ≥ 5.':
        'R = {r}: intervals are very wide, use R ≥ 5 for the report.',
    'середній':
        'mean',
    '99-й перцентиль':
        '99th percentile',
    'максимальний':
        'maximum',
    'дедлайн':
        'deadline',
    'немає даних про окремі події':
        'no per-event data',
    'Час реакції в часі (перший прогін, seed = %s)':
        'Response time over time (first replication, seed = %s)',
    'Увага: U = {u} ≥ 1 – процесор перевантажений, тому час реакції МРЧ постійно зростає. Після останньої події МРЧ модель лише дообробляє чергу.':
        'Note: U = {u} ≥ 1 – the processor is overloaded, so the SRT response time keeps growing. After the last SRT event the model only drains the queue.',
    'за весь час моделювання':
        'over the whole simulation',
    '⌂ Скинути':
        '⌂ Reset',
    'Повернути початковий вигляд графіка':
        'Restore the initial view of the chart',
    'Попередній вигляд':
        'Previous view',
    'Наступний вигляд':
        'Next view',
    'графік':
        'chart',
    'Графік збережено: {p}':
        'Chart saved: {p}',
    'Сценарій «{s}»':
        'Scenario “{s}”',
    'Немає даних для графіка':
        'No data for the chart',
    'Не вдалося побудувати графік: {e}':
        'Failed to build the chart: {e}',
    '«{a}» ({x}) і «{b}» ({y})':
        '“{a}” ({x}) and “{b}” ({y})',
    'Дедлайн МРЧ D1, мс':
        'SRT deadline D1, ms',
    ' і ':
        ' and ',
    'під час відмови':
        'during the fault',
    'Не вдалося зберегти графік: {e}':
        'Failed to save the chart: {e}',
    ', мс':
        ', ms',
    'Окремі лінії – для різних значень ':
        'Separate lines for different values of ',
    'Кожен стовпець – окремий сценарій з двома типами подій (МРЧ і ЖРЧ). Комірки можна редагувати прямо в таблиці.':
        'Each column is a separate scenario with two event types (SRT and HRT). Cells can be edited directly in the table.',
    'Файл не відкрито – можна заповнити таблицю вручну.':
        'No file open – you can fill in the table manually.',
    'U = λ1·(s1 + s2) + λ2·(s1 + s2b) – частка часу, коли процесор зайнятий':
        'U = λ1·(s1 + s2) + λ2·(s1 + s2b) – the fraction of time the processor is busy',
    'Зберегти таблицю параметрів':
        'Save parameter table',
    'Відкрити Excel…':
        'Open Excel…',
    'Зберегти як…':
        'Save as…',
    '▶  Запустити серію':
        '▶  Run series',
    '+ Стовпець':
        '+ Column',
    'Дублювати':
        'Duplicate',
    'Видалити':
        'Delete',
    'Червоним підсвічено значення поза допустимими межами (підказка – при наведенні). Рядок U – завантаження процесора: зелений < 0,95, жовтий 0,95–1, червоний ≥ 1 (перевантаження). Порожні стовпці не моделюються.':
        'Values outside the allowed range are highlighted in red (hint on hover). Row U is processor utilization: green < 0.95, yellow 0.95–1, red ≥ 1 (overload). Empty columns are not simulated.',
    'Файл: {f}  ·  сценаріїв: {n}':
        'File: {f}  ·  scenarios: {n}',
    'Таблицю заповнено вручну  ·  сценаріїв: {n}':
        'Table filled in manually  ·  scenarios: {n}',
    'Спочатку виділіть комірку в стовпці, який треба скопіювати.':
        'First select a cell in the column to copy.',
    'Виділіть комірки в стовпцях, які треба видалити.':
        'Select cells in the columns to delete.',
    'Видалити стовпці':
        'Delete columns',
    'Видалити сценарії {n}?':
        'Delete scenarios {n}?',
    'Відкрити таблицю параметрів':
        'Open parameter table',
    'Excel (*.xlsx *.xls);;Усі файли (*)':
        'Excel (*.xlsx *.xls);;All files (*)',
    'параметри.xlsx':
        'parameters.xlsx',
    'Таблицю збережено: {p}':
        'Table saved: {p}',
    'Завантажити таблицю параметрів з .xlsx':
        'Load a parameter table from .xlsx',
    'Додати порожній сценарій':
        'Add an empty scenario',
    'Скопіювати виділений стовпець у новий сценарій':
        'Copy the selected column into a new scenario',
    'Видалити виділені стовпці':
        'Delete the selected columns',
    'U (завантаження)':
        'U (utilization)',
    '«{t}» – не число. Використовуйте цифри, кома або крапка – десятковий роздільник.':
        '“{t}” is not a number. Use digits; comma or dot as the decimal separator.',
    'Не вдалося відкрити файл: {e}':
        'Failed to open the file: {e}',
    'Файл відкрито, але є зауваження:\n• ':
        'File opened, but there are remarks:\n• ',
    'Не вдалося зберегти: {e}':
        'Failed to save: {e}',
    'Процесор не перевантажений':
        'The processor is not overloaded',
    'Близько до перевантаження: потрібні більші N і R':
        'Close to overload: larger N and R are needed',
    'Перевантаження: черга МРЧ росте необмежено':
        'Overload: the SRT queue grows without bound',
    'Натисніть «Запустити серію».':
        'Press “Run series”.',
    'Відкрито «{f}»: {n} сценаріїв.':
        'Opened “{f}”: {n} scenarios.',
    'Збережено':
        'Saved',
    'Теги':
        'Tags',
    'Ідентифікатор':
        'Identifier',
    'Результати, збережені кнопкою «Зберегти в базу даних».':
        'Results saved with the “Save to database” button.',
    'Пошук за назвою сценарію або тегом (наприклад, назвою Excel-файлу)…':
        'Search by scenario name or tag (e.g. Excel file name)…',
    'Показано {a} з {b}. Подвійний клік – відкрити.':
        'Showing {a} of {b}. Double-click to open.',
    'Видалено: {n}.':
        'Deleted: {n}.',
    'Порівняння результатів':
        'Comparison of results',
    'Оновити':
        'Refresh',
    'Відкрити':
        'Open',
    'Порівняти':
        'Compare',
    'Сховище: {d}':
        'Storage: {d}',
    'Виділіть рядок у таблиці.':
        'Select a row in the table.',
    'Виділіть рядки, які треба видалити.':
        'Select the rows to delete.',
    'Видалити результати':
        'Delete results',
    'Видалити результатів зі сховища: {n}? Цю дію не можна скасувати.':
        'Delete results from storage: {n}? This cannot be undone.',
    'Для порівняння виділіть щонайменше два рядки (Ctrl + клік).':
        'Select at least two rows to compare (Ctrl + click).',
    'Середнє ± 95 % ДІ. Часи – у мілісекундах. Найкраще значення в рядку виділено жирним (для показників, де менше – краще).':
        'Mean ± 95 % CI. Times in milliseconds. The best value in a row is bold (for metrics where smaller is better).',
    'Закрити':
        'Close',
    'Показати вибраний результат на сторінці «Результати»':
        'Show the selected result on the “Results” page',
    'Виділіть 2 або більше рядків (Ctrl + клік)':
        'Select 2 or more rows (Ctrl + click)',
    'Не вдалося прочитати сховище: {e}':
        'Failed to read the storage: {e}',
    'Ще немає результатів. Запустіть серію з Excel або сценарій.':
        'No results yet. Run a series from Excel or a scenario.',
    'Показники':
        'Metrics',
    'Показувати ± 95 % довірчий інтервал':
        'Show ± 95 % confidence interval',
    'Під час дії відмов':
        'During faults',
    'Показники лише для подій, що надійшли, поки діяла відмова (без відкидання перших і останніх 5 % – відмова є перехідним процесом).':
        'Metrics only for events that arrived while the fault was active (without discarding the first and last 5 % – a fault is a transient process).',
    'Попередження і помилки':
        'Warnings and errors',
    'Результати серії':
        'Series results',
    'сценаріїв змодельовано':
        'scenarios simulated',
    'з попередженнями':
        'with warnings',
    'з помилками (не моделювались)':
        'with errors (not simulated)',
    'макс. P порушення дедлайну ЖРЧ':
        'max P deadline miss, HRT',
    "Часи – у мілісекундах. Жовтим виділено порушення м'яких дедлайнів (МРЧ), червоним – жорстких (ЖРЧ). ⚠ – сценарій з попередженням, ✖ – з помилкою (наведіть на заголовок).":
        'Times in milliseconds. Soft deadline misses (SRT) are highlighted in yellow, hard ones (HRT) in red. ⚠ – scenario with a warning, ✖ – with an error (hover over the header).',
    'Результати сценарію':
        'Scenario results',
    'завантаження U':
        'utilization U',
    'Часи – у мілісекундах; час реакції = від надходження переривання до завершення обробки. Червоним виділено порушення дедлайнів ЖРЧ, жовтим – МРЧ.':
        'Times in milliseconds; response time = from the interrupt arrival to the end of processing. HRT deadline misses are red, SRT – yellow.',
    'Результат з бази даних':
        'Result from the database',
    'P порушення дедлайну ЖРЧ':
        'P deadline miss, HRT',
    'P порушення дедлайну МРЧ':
        'P deadline miss, SRT',
    'версія програми':
        'program version',
    'Відкрито зі сховища. Щоб перерахувати, відкрийте сценарій і запустіть знову.':
        'Opened from storage. To recompute, open the scenario and run it again.',
    'Збережено в {d}: нових результатів – {n}':
        'Saved to {d}: new results – {n}',
    'сценарій':
        'scenario',
    'Експорт результатів':
        'Export results',
    'Зберегти в базу даних':
        'Save to database',
    'Експорт в Excel…':
        'Export to Excel…',
    'Графіки →':
        'Charts →',
    'розрахунок {t} с':
        'computed in {t} s',
    'подій ЖРЧ у сценарії немає':
        'no HRT events in the scenario',
    'макс. P порушення дедлайну МРЧ':
        'max P deadline miss, SRT',
    'подій МРЧ у сценарії немає':
        'no SRT events in the scenario',
    'немає':
        'none',
    'відмов у сценарії':
        'faults in the scenario',
    'Дедлайн {d} мс, WCET + T_core = {w} мс, інтенсивність {r} 1/мс':
        'Deadline {d} ms, WCET + T_core = {w} ms, rate {r} 1/ms',
    'збережено {d}  ·  теги: {t}':
        'saved {d}  ·  tags: {t}',
    '; вже були в базі (ті самі параметри, R і seed), не дублювались – {n}':
        '; already in the database (same parameters, R and seed), not duplicated – {n}',
    'Звіт збережено: {p}':
        'Report saved: {p}',
    'Записати сценарії і результати в MongoDB. Повторне збереження не створює дублікатів.':
        'Store scenarios and results in MongoDB. Saving again does not create duplicates.',
    'Помилка: ':
        'Error: ',
    'Без зауважень':
        'No remarks',
    'хибних переривань ≈ {n}':
        'spurious interrupts ≈ {n}',
    'подій з перевищенням WCET ≈ {n}':
        'events with WCET overrun ≈ {n}',
    'Під час дії відмов: ':
        'During faults: ',
    'Не вдалося зберегти в базу даних: {e}':
        'Failed to save to the database: {e}',
    'Не вдалося експортувати: {e}':
        'Failed to export: {e}',
    'сценарій {n}':
        'scenario {n}',
    'Назва':
        'Name',
    'Як надходять':
        'Arrivals',
    'WCET, мс':
        'WCET, ms',
    'Фаза, мс':
        'Phase, ms',
    'Назва типу подій (унікальна)':
        'Event type name (unique)',
    "ЖРЧ – жорсткий реальний час (вищий пріоритет), МРЧ – м'який":
        'HRT – hard real-time (higher priority), SRT – soft',
    'Закон надходження: за замовчуванням ЖРЧ – періодичний, МРЧ – пуассонівський':
        'Arrival law: by default HRT is periodic, SRT is Poisson',
    'Скільки подій надходить за 1 мс (для періодичних – 1/період)':
        'How many events arrive per 1 ms (for periodic ones – 1/period)',
    'Тривалість вторинної обробки події (без T_core)':
        'Secondary processing time of an event (without T_core)',
    'Допустимий час реакції на подію':
        'Allowed response time for an event',
    'Момент першої періодичної події (порожньо – через один період)':
        'Moment of the first periodic event (empty – after one period)',
    'Увімк.':
        'On',
    'Вид відмови':
        'Fault kind',
    'Початок, мс':
        'Start, ms',
    'Тривалість, мс':
        'Duration, ms',
    'Шторм: переривань/мс':
        'Storm: interrupts/ms',
    'Шторм: обробка, мс':
        'Storm: processing, ms',
    'WCET: тип подій':
        'WCET: event type',
    'WCET: імовірність':
        'WCET: probability',
    'WCET: множник':
        'WCET: factor',
    'Вимкнена відмова зберігається у файлі, але не моделюється':
        'A disabled fault is kept in the file but not simulated',
    'Шторм переривань – хибні переривання; Перевищення WCET – обробка триває довше':
        'Interrupt storm – spurious interrupts; WCET overrun – processing takes longer',
    'Довільна назва для звіту':
        'Any name for the report',
    'Момент модельного часу, з якого діє відмова':
        'Model time from which the fault is active',
    'Скільки діє відмова (порожньо – до кінця моделювання)':
        'How long the fault lasts (empty – until the end of the simulation)',
    'Інтенсивність хибних переривань під час шторму':
        'Rate of spurious interrupts during the storm',
    'Вторинна обробка одного хибного переривання на рівні МРЧ (0 – лише ISR)':
        'Secondary processing of one spurious interrupt at SRT level (0 – ISR only)',
    'Тип подій, обробка яких затягується (або «усі типи»)':
        'Event type whose processing is delayed (or “all types”)',
    'Частка подій з перевищенням (0…1)':
        'Fraction of events with an overrun (0…1)',
    'У скільки разів обробка довша за WCET':
        'How many times longer than WCET the processing takes',
    'Шторм переривань':
        'Interrupt storm',
    'Перевищення WCET':
        'WCET overrun',
    'Кілька типів подій і відмови':
        'Several event types and faults',
    'Нова можливість (побажання керівника): не 2 типи подій, як в Excel, а скільки завгодно, і сценарії відмов. Для роботи з Excel-файлами цей розділ не потрібен.':
        "New feature (supervisor's request): not 2 event types as in Excel but any number, plus fault scenarios. This section is not needed for Excel files.",
    'Готові сценарії з папки scenarios':
        'Ready-made scenarios from the scenarios folder',
    "<b>Як користуватися:</b> 1) угорі праворуч виберіть готовий сценарій (або «Новий»); 2) за потреби змініть числа в таблиці «Типи подій» і додайте відмову кнопкою «Шторм переривань» або «Перевищення WCET»; 3) натисніть «Запустити» – відкриються «Результати», далі «Графіки». Наведіть мишу на заголовок стовпця – з'явиться пояснення.":
        '<b>How to use:</b> 1) at the top right choose a ready-made scenario (or “New”); 2) if needed, change the numbers in the “Event types” table and add a fault with “Interrupt storm” or “WCET overrun”; 3) press “Run” – “Results” opens, then “Charts”. Hover over a column header to see an explanation.',
    'Загальне':
        'General',
    'Короткий опис: звідки взято значення':
        'Short description: where the values come from',
    'Кількість подій МРЧ, що моделюються (якщо МРЧ немає – усіх подій)':
        'Number of SRT events to simulate (if there are no SRT events – all events)',
    'Частка подій, що відкидається на початку і в кінці (стаціонарність)':
        'Fraction of events discarded at the beginning and the end (steady state)',
    'Опис':
        'Description',
    'N (подій МРЧ)':
        'N (SRT events)',
    'Частка розгону':
        'Warm-up fraction',
    'Платформа':
        'Platform',
    'Задайте частоту і такти – час буде перераховано за T = C / F. Або залиште такти порожніми і задайте час у мс.':
        'Set the frequency and cycles – times are computed as T = C / F. Or leave the cycles empty and set the times in ms.',
    'Типи подій':
        'Event types',
    'ЖРЧ обробляються з вищим пріоритетом (черга QH), МРЧ – з нижчим (QS). Будь-яке нове переривання перериває вторинну обробку.':
        'HRT events are processed with higher priority (queue QH), SRT – with lower (QS). Any new interrupt preempts secondary processing.',
    'Відмови':
        'Faults',
    'Що станеться, якщо щось піде не так. Шторм переривань – хибні переривання (наприклад, від несправної лінії). Перевищення WCET – обробка триває довше, ніж заплановано. Відмова діє лише на заданому інтервалі часу.':
        'What happens if something goes wrong. Interrupt storm – spurious interrupts (e.g. from a faulty line). WCET overrun – processing takes longer than planned. A fault is active only in the given time interval.',
    'Відмов немає – система працює в нормальному режимі.':
        'No faults – the system works normally.',
    'Підсумок перед запуском':
        'Summary before running',
    '— готові сценарії —':
        '— ready-made scenarios —',
    'Відкрито: {f}':
        'Opened: {f}',
    "ЖРЧ – жорсткий реальний час: пропуск дедлайну неприпустимий, обробляється першою.\nМРЧ – м'який реальний час: запізнення небажане, але допустиме.":
        'HRT – hard real-time: a deadline miss is unacceptable, processed first.\nSRT – soft real-time: lateness is undesirable but acceptable.',
    'автоматично':
        'automatic',
    'періодично':
        'periodic',
    'випадково':
        'random',
    'Як надходять події цього типу:\n• автоматично – ЖРЧ періодично, МРЧ випадково (так у більшості задач);\n• періодично – рівно через 1/інтенсивність мс (датчики, таймери);\n• випадково – потік Пуассона, інтервали випадкові (команди, повідомлення).':
        'How events of this type arrive:\n• automatic – HRT periodic, SRT random (as in most tasks);\n• periodic – exactly every 1/rate ms (sensors, timers);\n• random – Poisson flow, random intervals (commands, messages).',
    'Створено новий сценарій з двома типами подій – змініть його під свою задачу.':
        'A new scenario with two event types has been created – adapt it to your task.',
    'Сценарій (*.json)':
        'Scenario (*.json)',
    'Зберегти сценарій':
        'Save scenario',
    'Відкрити…':
        'Open…',
    'Новий':
        'New',
    '▶  Запустити':
        '▶  Run',
    'F_clk, МГц':
        'F_clk, MHz',
    'Тактова частота процесора':
        'Processor clock frequency',
    'C_IRQ, тактів':
        'C_IRQ, cycles',
    'Такти первинного обробника (ISR)':
        'Primary handler (ISR) cycles',
    'C_core, тактів':
        'C_core, cycles',
    'Такти накладних витрат ОС на одну подію':
        'OS overhead cycles per event',
    'T_IRQ, мс':
        'T_IRQ, ms',
    'Час ISR напряму (якщо такти не задано)':
        'ISR time directly (if cycles are not set)',
    'T_core, мс':
        'T_core, ms',
    'Накладні витрати ОС напряму (якщо такти не задано)':
        'OS overhead directly (if cycles are not set)',
    '+ Тип подій':
        '+ Event type',
    '+ Шторм переривань':
        '+ Interrupt storm',
    '+ Перевищення WCET':
        '+ WCET overrun',
    'Завантаження процесора:':
        'Processor utilization:',
    'усі типи':
        'all types',
    'Тип {n}':
        'Type {n}',
    '«{n}»: інтенсивність':
        '“{n}”: rate',
    '«{n}»: дедлайн':
        '“{n}”: deadline',
    '«{n}»: фаза':
        '“{n}”: phase',
    'Відмова {n}':
        'Fault {n}',
    '«{n}»: початок':
        '“{n}”: start',
    '«{n}»: тривалість':
        '“{n}”: duration',
    'Новий сценарій':
        'New scenario',
    'новий_сценарій.json':
        'new_scenario.json',
    'Сценарій збережено: {p}':
        'Scenario saved: {p}',
    'T_IRQ = {a} мс,  T_core = {b} мс':
        'T_IRQ = {a} ms,  T_core = {b} ms',
    'Первинна обробка T_IRQ = {a} мс':
        'Primary processing T_IRQ = {a} ms',
    'Не вдалося відкрити сценарій: {e}':
        'Failed to open the scenario: {e}',
    '{w}: не заповнено':
        '{w}: not filled',
    '{w}: «{t}» – не число':
        '{w}: “{t}” is not a number',
    '{w}: некоректне число':
        '{w}: invalid number',
    'БО':
        'OBC',
    'Спочатку виправте помилки:\n• ':
        'Fix the errors first:\n• ',
    'Сценарій не можна запустити:\n• ':
        'The scenario cannot be run:\n• ',
    'Задайте T_IRQ (у мс) або F_clk і C_IRQ.':
        'Set T_IRQ (in ms) or F_clk and C_IRQ.',
    'під час відмов U = {u}':
        'during faults U = {u}',
    '«{n}»: інтенсивність шторму':
        '“{n}”: storm rate',
    '«{n}»: обробка':
        '“{n}”: processing',
    '«{n}»: імовірність':
        '“{n}”: probability',
    '«{n}»: множник':
        '“{n}”: factor',
    'Сценарій коректний, можна запускати.':
        'The scenario is correct, ready to run.',
    'Середній час очікування первинної обробки':
        'Mean waiting time for primary processing',
    "Середній час очікування вторинної обробки для подій м'якого РЧ":
        'Mean waiting time for secondary processing of soft real-time events',
    'Середній час очікування вторинної обробки для подій жорсткого РЧ':
        'Mean waiting time for secondary processing of hard real-time events',
    "Середній повний час перебування події м'якого РЧ у системі":
        'Mean total sojourn time of soft real-time events in the system',
    'Середній повний час перебування подій жорсткого РЧ у системі':
        'Mean total sojourn time of hard real-time events in the system',
    'Максимальний час очікування первинної обробки':
        'Maximum waiting time for primary processing',
    "Максимальний час очікування вторинної обробки подій м'якого РЧ":
        'Maximum waiting time for secondary processing of soft real-time events',
    'Максимальний час очікування вторинної обробки подій жорсткого РЧ':
        'Maximum waiting time for secondary processing of hard real-time events',
    "Максимальний повний час перебування подій м'якого РЧ у системі":
        'Maximum total sojourn time of soft real-time events in the system',
    'Максимальний повний час перебування подій жорсткого РЧ у системі':
        'Maximum total sojourn time of hard real-time events in the system',
    'Середня довжина черги первинного обробника':
        'Mean queue length of the primary handler',
    "Середня довжина черги вторинного обробника для подій м'якого РЧ":
        'Mean queue length of the secondary handler for soft real-time events',
    'Середня довжина черги вторинного обробника для подій жорсткого РЧ':
        'Mean queue length of the secondary handler for hard real-time events',
    'Коефіцієнт завантаження первинного обробника':
        'Utilization of the primary handler',
    "Коефіцієнт завантаження вторинного обробника подіями м'якого РЧ":
        'Utilization of the secondary handler by soft real-time events',
    'Коефіцієнт завантаження вторинного обробника подіями жорсткого РЧ':
        'Utilization of the secondary handler by hard real-time events',
    "Імовірність порушення дедлайну подіями м'якого РЧ":
        'Deadline miss probability for soft real-time events',
    'Імовірність порушення дедлайну подіями жорсткого РЧ':
        'Deadline miss probability for hard real-time events',
    'λ1, мс⁻¹ (інтенсивність потоку подій МРЧ)':
        'λ1, ms⁻¹ (rate of the SRT event flow)',
    's1, мс (тривалість первинної обробки)':
        's1, ms (primary processing time)',
    's2, мс (тривалість вторинної обробки подій МРЧ)':
        's2, ms (secondary processing time of SRT events)',
    'N (кількість подій МРЧ для моделювання)':
        'N (number of SRT events to simulate)',
    'D1, мс (дедлайн МРЧ)':
        'D1, ms (SRT deadline)',
    'λ2, мс⁻¹ (частота подій ЖРЧ)':
        'λ2, ms⁻¹ (frequency of HRT events)',
    's2b, мс (тривалість вторинної обробки подій ЖРЧ)':
        's2b, ms (secondary processing time of HRT events)',
    'D2, мс (дедлайн ЖРЧ)':
        'D2, ms (HRT deadline)',
    'λ2, мс⁻¹ (частота подій ЖРЧ; 0 – потоку ЖРЧ немає)':
        'λ2, ms⁻¹ (frequency of HRT events; 0 – no HRT flow)',
    # ---- версія 4.3: стартовий екран, «Про програму», очищення таблиці ----
    '1 сценарій':
        '1 scenario',
    '{n} сценаріїв':
        '{n} scenarios',
    'Переривання':
        'Interrupts',
    'ЖРЧ · періодично':
        'HRT · periodic',
    'МРЧ · випадково':
        'SRT · random',
    'первинна\nобробка, Q0':
        'primary\nprocessing, Q0',
    'QH · ЖРЧ – вищий пріоритет':
        'QH · HRT – higher priority',
    'QS · МРЧ – нижчий пріоритет':
        'QS · SRT – lower priority',
    'Процесор':
        'Processor',
    'вторинна\nобробка (WCET)':
        'secondary\nprocessing (WCET)',
    'нова подія перериває\nвторинну обробку':
        'a new event preempts\nsecondary processing',
    'Серія сценаріїв':
        'Scenario series',
    'Про програму':
        'About',
    'Типи подій і відмови':
        'Event types & faults',
    'Очистити таблицю':
        'Clear table',
    'Темна тема':
        'Dark theme',
    'Світла тема':
        'Light theme',
    'Початок':
        'Home',
    'Історія':
        'History',
    'Розрахунок показників реактивності бортового обчислювача: час реакції на переривання, черги, завантаження процесора і ймовірність порушення дедлайнів. Модель – три черги з пріоритетами і витісненням, імітаційне моделювання з довірчими інтервалами.':
        'Calculation of on-board computer reactivity indicators: interrupt response time, queues, processor utilization and deadline miss probability. The model has three queues with priorities and preemption; discrete-event simulation with confidence intervals.',
    'таблиця параметрів, 2 типи подій':
        'parameter table, 2 event types',
    'довільні типи подій, шторм, WCET':
        'any event types, storm, WCET',
    'збережені запуски, порівняння':
        'saved runs, comparison',
    'Схема моделі: події проходять первинну обробку (ISR), потім чекають у черзі свого класу; ЖРЧ обслуговуються раніше за МРЧ':
        'Model diagram: events pass primary processing (ISR), then wait in the queue of their class; HRT events are served before SRT',
    'Задайте параметри: відкрийте Excel-файл, готовий сценарій або введіть значення вручну.':
        'Set the parameters: open an Excel file, a ready-made scenario, or enter values manually.',
    'Ліворуч унизу оберіть кількість прогонів R (для звіту – 5 і більше) і seed.':
        'At the bottom left choose the number of replications R (5 or more for a report) and the seed.',
    'Натисніть «Запустити» – розрахунок іде у фоні на всіх ядрах процесора.':
        'Press “Run” – the computation runs in the background on all processor cores.',
    'Перегляньте «Результати» і «Графіки», збережіть звіт в Excel або результат у базу.':
        'Review “Results” and “Charts”, export a report to Excel or save the result to the database.',
    'Останні результати':
        'Recent results',
    'Подвійний клік – відкрити в історії':
        'Double-click – open in history',
    'Реактивність БО':
        'OBC Reactivity',
    'Магістерська робота, {y}':
        "Master's thesis, {y}",
    'Науковий керівник: {s}':
        'Supervisor: {s}',
    'Детальніше':
        'Details',
    'ХАІ':
        'KhAI',
    'Таблиця параметрів: кожен стовпець – окремий сценарій з двома типами подій (МРЧ і ЖРЧ). Відкрийте Excel-файл або введіть значення вручну.':
        'Parameter table: each column is a separate scenario with two event types (SRT and HRT). Open an Excel file or enter values manually.',
    'Запустити серію':
        'Run series',
    'Стовпець':
        'Column',
    'Повернути порожню таблицю, як під час запуску програми':
        'Reset to an empty table, as at program start',
    'Запустити':
        'Run',
    'Тип подій':
        'Event type',
    'Ще немає збережених результатів. Після запуску натисніть «Зберегти в базу даних».':
        'No saved results yet. After a run press “Save to database”.',
    'Версія {v} · модель {m}':
        'Version {v} · model {m}',
    'Програма розраховує показники реактивності бортового обчислювача методом імітаційного моделювання: час реакції на переривання, довжини черг, завантаження процесора і ймовірність порушення дедлайнів, у тому числі під час відмов.':
        'The program calculates on-board computer reactivity indicators by discrete-event simulation: interrupt response time, queue lengths, processor utilization and deadline miss probability, including during faults.',
    'Розробник':
        'Developer',
    'Заклад':
        'Institution',
    'Науковий керівник':
        'Supervisor',
    'Тема роботи':
        'Thesis topic',
    'Пошта':
        'E-mail',
    'Технології':
        'Technologies',
    'Дані розробника можна змінити у файлі config.py (розділ DEVELOPER).':
        'Developer details can be changed in config.py (DEVELOPER section).',
    'Очистити всю таблицю? Незбережені зміни буде втрачено.':
        'Clear the whole table? Unsaved changes will be lost.',
    'Таблицю очищено. Відкрийте Excel-файл або введіть значення вручну.':
        'The table has been cleared. Open an Excel file or enter values manually.',
    # --- пояснення простими словами (v4.4) ---
    'випадково, в середньому раз на {t} мс':
        'at random, on average once every {t} ms',
    'запас продуктивності достатній':
        'there is enough spare capacity',
    "Процесор зайнятий у середньому на <b style='color:{c}'>{u}</b> часу – {w}.":
        "On average the processor is busy <b style='color:{c}'>{u}</b> of the time – {w}.",
    'Запізнюється значна частина подій – бракує продуктивності процесора або дедлайн надто жорсткий.':
        'A large share of events is late – the processor lacks capacity or the deadline is too tight.',
    '{when} в {p} випадків обробка {t} триває <b>в {k} раза довше</b>, ніж заплановано (перевищення WCET – наприклад, через повтори обміну по шині).':
        '{when}, in {p} of cases processing of {t} takes <b>{k} times longer</b> than planned (WCET overrun – e.g. due to bus retries).',
    '< {v} %':
        '< {v} %',
    'кожні {t} мс':
        'every {t} ms',
    'процесор перевантажений – він не встигає обробляти події, черги ростуть без кінця':
        'the processor is overloaded – it cannot keep up, queues grow without bound',
    'процесор майже повністю зайнятий – черги можуть бути дуже довгими':
        'the processor is almost fully busy – queues can become very long',
    'запас продуктивності невеликий':
        'there is little spare capacity',
    "Процесору потрібно <b style='color:{c}'>{u}</b> його потужності – більше, ніж він має: {w}.":
        "The processor would need <b style='color:{c}'>{u}</b> of its capacity – more than it has: {w}.",
    "Частка мала – для м'яких подій це зазвичай допустимо.":
        'The share is small – usually acceptable for soft events.',
    'Частка помітна – варто оцінити, чи це прийнятно.':
        'The share is noticeable – check whether it is acceptable.',
    'З {a} по {b} с модельного часу':
        'From {a} to {b} s of model time',
    'Починаючи з {a} с модельного часу':
        'Starting from {a} s of model time',
    "{when} надходять <b>хибні переривання</b> (наприклад, від несправної лінії зв'язку): {r} на мілісекунду, кожне забирає {w} мс процесора.":
        '{when}, <b>spurious interrupts</b> arrive (e.g. from a faulty communication line): {r} per millisecond, each taking {w} ms of processor time.',
    'усіх типів подій':
        'all event types',
    'Змодельовано сценаріїв: {n} із {m}.':
        'Scenarios simulated: {n} of {m}.',
    'Щоб побачити, як показники змінюються від параметра, відкрийте «Графіки».':
        'To see how the metrics depend on a parameter, open “Charts”.',
    'бортовий обчислювач':
        'on-board computer',
    '<b>Що моделюється.</b> Обчислювач {p} отримує переривання від {n} джерел подій. Кожна подія спочатку проходить коротку первинну обробку (ISR), потім чекає в черзі свого класу.':
        '<b>What is simulated.</b> The computer {p} receives interrupts from {n} event sources. Each event first gets short primary processing (ISR), then waits in the queue of its class.',
    'Буде змодельовано приблизно {t} с роботи системи.':
        'About {t} s of system operation will be simulated.',
    '<b>Відмов немає</b> – система працює в нормальному режимі.':
        '<b>No faults</b> – the system operates normally.',
    "<b>На що дивитися після запуску.</b> У «Результатах» з'явиться окрема таблиця «Під час дії відмов» – порівняйте її з основною. Щоб чітко побачити вплив відмови, запустіть цей самий сценарій ще раз без неї (зніміть галочку в таблиці «Відмови») і порівняйте обидва запуски в «Історії».":
        '<b>What to look at after the run.</b> “Results” will show a separate table “During faults” – compare it with the main one. To see the effect of the fault clearly, run the same scenario again without it (untick it in the “Faults” table) and compare both runs in “History”.',
    '<b>На що дивитися після запуску.</b> Головне – «P порушення дедлайну»: для ЖРЧ вона має бути 0. Далі – найдовший час реакції кожного типу порівняно з його дедлайном.':
        '<b>What to look at after the run.</b> The key metric is “P of deadline miss”: for HRT it must be 0. Next, compare the longest response time of each type with its deadline.',
    'З попередженнями (⚠ у заголовку – наведіть мишу): {s}.':
        'With warnings (⚠ in the header – hover over it): {s}.',
    'Пропуск їхнього дедлайну – небезпечна відмова.':
        'Missing their deadline is a dangerous failure.',
    'Запізнення допустимі, але небажані.':
        'Delays are tolerable but undesirable.',
    'Найповільніша реакція – у сценарії {c}: у середньому {a} мс для МРЧ і {b} мс для ЖРЧ.':
        'Slowest response – scenario {c}: on average {a} ms for SRT and {b} ms for HRT.',
    '<b>Відмова.</b>':
        '<b>Fault.</b>',
    "Поки діє відмова, процесору потрібно <b style='color:{c}'>{u}</b> його потужності – більше, ніж він має, тому черги ростимуть, поки відмова не закінчиться.":
        "While the fault lasts the processor would need <b style='color:{c}'>{u}</b> of its capacity – more than it has, so queues will grow until the fault ends.",
    "Поки діє відмова, процесор зайнятий на <b style='color:{c}'>{u}</b> часу.":
        "While the fault lasts the processor is busy <b style='color:{c}'>{u}</b> of the time.",
    'Вимкнені відмови (без галочки) не враховуються: {n}.':
        'Disabled faults (unticked) are ignored: {n}.',
    'Найближче до межі – {n}: найдовша реакція {t} мс при дедлайні {d} мс.':
        'Closest to the limit – {n}: longest response {t} ms with a deadline of {d} ms.',
    '{n} – {p} (за весь час {q})':
        '{n} – {p} (over the whole run {q})',
    "Жорсткі події не постраждали – їх захищає вищий пріоритет; весь удар прийняли м'які події.":
        'Hard events were not affected – the higher priority protects them; soft events took the whole hit.',
    'Не змодельовано через помилки в даних: {s}.':
        'Not simulated because of data errors: {s}.',
    'Жорсткі (ЖРЧ), вищий пріоритет':
        'Hard (HRT), higher priority',
    "М'які (МРЧ), нижчий пріоритет":
        'Soft (SRT), lower priority',
    'Для бортової системи це неприпустимо.':
        'This is unacceptable for an on-board system.',
    "М'які події (МРЧ) теж укладаються в дедлайни.":
        'Soft events (SRT) also meet their deadlines.',
    'Відмова ({f}) не призвела до порушень дедлайнів – система її витримала.':
        'The fault ({f}) caused no deadline misses – the system withstood it.',
    'Відмова зачепила жорсткі події – пріоритет від цього не захищає, бо збій стався в обробці самої жорсткої задачі або процесор перевантажений.':
        'The fault affected hard events – priority does not protect against this, because the failure is in the hard task itself or the processor is overloaded.',
    'Жорсткі дедлайни (ЖРЧ) порушуються у сценаріях: {s}.':
        'Hard deadlines (HRT) are missed in scenarios: {s}.',
    'У жодному сценарії жорсткі дедлайни (ЖРЧ) не порушуються.':
        'Hard deadlines (HRT) are met in every scenario.',
    "М'які дедлайни (МРЧ) порушуються у сценаріях: {s}.":
        'Soft deadlines (SRT) are missed in scenarios: {s}.',
    "М'які дедлайни (МРЧ) теж виконуються всюди.":
        'Soft deadlines (SRT) are met everywhere too.',
    '⚠ Відмова почнеться вже після кінця моделювання (≈ {t} с) – збільште N.':
        '⚠ The fault starts after the simulation ends (≈ {t} s) – increase N.',
    'Усі жорсткі події (ЖРЧ) обробляються вчасно.':
        'All hard events (HRT) are processed on time.',
    '<b>Під час відмови</b> ({f}) запізнюються:':
        '<b>During the fault</b> ({f}) late events:',
    '⚠ Моделювання закінчиться раніше, ніж відмова (≈ {t} с).':
        '⚠ The simulation ends before the fault does (≈ {t} s).',
    "М'які події (МРЧ) запізнюються:":
        'Soft events (SRT) are late:',
    '{n} – у {p} випадків':
        '{n} – in {p} of cases',
    'Жорсткі дедлайни порушуються:':
        'Hard deadlines are missed:',
    'Тут систему описують детальніше, ніж у таблиці параметрів: скільки завгодно типів подій (а не лише 2) і можливі відмови – щоб перевірити, що станеться, коли щось піде не так.':
        'Here the system is described in more detail than in the parameter table: any number of event types (not just 2) and possible faults – to check what happens when something goes wrong.',
    'Що відбувається в цьому сценарії':
        'What happens in this scenario',
    '— виберіть готовий сценарій —':
        '— choose a ready-made scenario —',
    'Без відмов':
        'Without faults',
    'З відмовами (щоб порівняти з нормальною роботою)':
        'With faults (to compare with normal operation)',
    "Виправте помилки в даних (див. «Підсумок перед запуском» унизу) – тоді тут з'явиться пояснення сценарію.":
        'Fix the data errors (see “Summary before run” below) – then an explanation of the scenario will appear here.',
    'Висновок простими словами':
        'Conclusion in plain words',
    'Боривітер: 6 типів подій, без відмов':
        'Boryviter: 6 event types, no faults',
    "Бортовий обчислювач супутника «Боривітер» у нормальній роботі: 3 типи жорстких подій (орієнтація ADCS, команди приводам) і 3 типи м'яких (команди з Землі, моніторинг живлення, телеметрія). Значення ілюстративні (на основі розділу 1), не виміряні. Для датчиків ADCS задано обмеження «4 з 5»: з будь-яких 5 опитувань поспіль хоча б 4 – вчасно.":
        'On-board computer of the “Boryviter” satellite in normal operation: 3 hard event types (ADCS attitude control, actuator commands) and 3 soft ones (ground commands, power monitoring, telemetry). Values are illustrative (based on chapter 1), not measured. ADCS sensors have the constraint “4 of 5”: in any 5 consecutive polls at least 4 are on time.',
    'Боривітер + відмова: повільні датчики ADCS':
        'Boryviter + fault: slow ADCS sensors',
    'Той самий «Боривітер», але з 20-ї по 40-ву секунду 30 % опитувань датчиків ADCS тривають у 2,5 раза довше (повтори обміну по шині I2C). Питання: що буде, коли збій стається в самій жорсткій задачі. Значення ілюстративні. Для датчиків ADCS задано обмеження «4 з 5»: з будь-яких 5 опитувань поспіль хоча б 4 – вчасно.':
        'The same “Boryviter”, but from 20 to 40 s, 30 % of ADCS sensor polls take 2.5 times longer (I2C bus retries). Question: what happens when the failure is in a hard task itself. Values are illustrative. ADCS sensors have the constraint “4 of 5”: in any 5 consecutive polls at least 4 are on time.',
    'Боривітер + відмова: шторм переривань UART':
        'Boryviter + fault: UART interrupt storm',
    'Той самий «Боривітер», але з 20-ї по 25-ту секунду несправна лінія UART видає хибні переривання (20 на мс, по 0,03 мс кожне). Питання: чи захистить пріоритет жорсткі задачі. Значення ілюстративні. Для датчиків ADCS задано обмеження «4 з 5»: з будь-яких 5 опитувань поспіль хоча б 4 – вчасно.':
        'The same “Boryviter”, but from 20 to 25 s a faulty UART line produces spurious interrupts (20 per ms, 0.03 ms each). Question: will priority protect the hard tasks. Values are illustrative. ADCS sensors have the constraint “4 of 5”: in any 5 consecutive polls at least 4 are on time.',
    'Базовий сценарій статті: 2 типи подій':
        'Baseline scenario from the paper: 2 event types',
    "Найпростіша модель зі статті: один тип жорстких і один тип м'яких подій (s2 = 0,5 мс) – те саме, що один стовпець таблиці параметрів.":
        'The simplest model from the paper: one hard and one soft event type (s2 = 0.5 ms) – the same as one column of the parameter table.',
    'Відмова ({f}) не встигла початися – моделювання закінчилося раніше. Збільште N, щоб модельний час охопив інтервал відмови.':
        'The fault ({f}) never started – the simulation ended earlier. Increase N so that model time covers the fault interval.',
    ' (частка запізнень {p})':
        ' (share of late events {p})',
    ' (частка запізнень від {a} до {b})':
        ' (share of late events from {a} to {b})',
    'у всіх сценаріях':
        'in all scenarios',
    'у сценаріях {s}':
        'in scenarios {s}',
    'Жорсткі дедлайни (ЖРЧ) порушуються {w}':
        'Hard deadlines (HRT) are missed {w}',
    "М'які дедлайни (МРЧ) порушуються {w}":
        'Soft deadlines (SRT) are missed {w}',
    'у всіх змодельованих сценаріях':
        'in all simulated scenarios',
    'повільна обробка {t}':
        'slow processing of {t}',
    'шторм переривань':
        'interrupt storm',
    'перевантаження':
        'overloaded',
    'Після запуску порівняйте таблицю «Під час дії відмов» з основною. Щоб побачити вплив відмови, запустіть сценарій ще раз без неї (зніміть галочку) і порівняйте обидва запуски в «Історії».':
        'After the run, compare the “During faults” table with the main one. To see the effect of the fault, run the scenario again without it (untick it) and compare both runs in “History”.',
    'Після запуску головне – «P порушення дедлайну»: для ЖРЧ вона має бути 0.':
        'After the run the key metric is “P of deadline miss”: for HRT it must be 0.',
    'типів подій: {h} ЖРЧ, {s} МРЧ':
        'event types: {h} HRT, {s} SRT',
    'майже на межі':
        'almost at the limit',
    'завантаження процесора – {w}':
        'processor load – {w}',
    'від {a} с':
        'from {a} s',
    '{n} відмови':
        '{n} faults',
    'модельного часу ({n} подій МРЧ)':
        'of model time ({n} SRT events)',
    'усіх типів':
        'all types',
    'запас невеликий':
        'little headroom',
    'запас є':
        'enough headroom',
    'під час відмови ({w}, {t})':
        'during the fault ({w}, {t})',
    'відмов – нормальна робота':
        'faults – normal operation',
    'Відмова {f} почнеться вже після кінця моделювання (≈ {t} с) – збільште N.':
        'Fault {f} starts after the simulation ends (≈ {t} s) – increase N.',
    'Моделювання закінчиться раніше, ніж відмова {f} (≈ {t} с).':
        'The simulation ends before fault {f} does (≈ {t} s).',
    'Система з будь-якою кількістю типів подій і з відмовами. Виберіть готовий сценарій (праворуч) → за потреби змініть числа → «Запустити». Підказки – на заголовках стовпців.':
        'A system with any number of event types and with faults. Choose a ready-made scenario (on the right) → change the numbers if needed → “Run”. Hints are on the column headers.',
    'Коротко про сценарій':
        'Scenario at a glance',
    'Під час відмови процесор не встигає – черги ростимуть, поки вона діє.':
        'During the fault the processor cannot keep up – queues will grow while it lasts.',
    'Виберіть сценарій → за потреби змініть числа → «Запустити».':
        'Choose a scenario → change the numbers if needed → “Run”.',
    'під час відмови: {w}, {t}':
        'during the fault: {w}, {t}',
    'шторм':
        'storm',
    # --- v4.5: (m, n), додаткові показники, порівняння з теорією ---
    '«{m} з {n}»':
        '“{m} of {n}”',
    '{t}: обмеження {r} виконується – у будь-яких {n} подіях поспіль пропусків не більше {k}.':
        '{t}: constraint {r} holds – any {n} consecutive events have at most {k} misses.',
    '{t}: обмеження {r} порушується у {p} вікон (у найгіршому вікні {w} пропусків з {n}, допустимо {k}).':
        '{t}: constraint {r} is violated in {p} of windows (worst window: {w} misses of {n}, allowed {k}).',
    'm з n':
        'm of n',
    'Обмеження weakly hard (m, n): з будь-яких n подій поспіль хоча б m мають бути оброблені вчасно.\nНаприклад, «4 з 5» – допускається не більше 1 пропуску дедлайну на 5 подій. Порожньо – без обмеження.':
        'Weakly hard constraint (m, n): in any n consecutive events at least m must be processed on time.\nFor example, “4 of 5” allows at most 1 deadline miss per 5 events. Empty – no constraint.',
    '{m} з {n}':
        '{m} of {n}',
    '«{n}»: WCET':
        '“{n}”: WCET',
    '«{n}»: m з n – введіть два цілі числа, наприклад «4 з 5»':
        '“{n}”: m of n – enter two whole numbers, e.g. “4 of 5”',
    'Системні показники':
        'System metrics',
    'U, H, U_os і η – за формулами розділу 1; максимальні довжини черг – за моделюванням.':
        'U, H, U_os and η – from the chapter 1 formulas; maximum queue lengths – from the simulation.',
    'Додаткові показники':
        'Additional metrics',
    'Показники розділу 1, яких немає серед 18 основних. «Мін. дедлайн для P ≤ 5 %» – найменший дедлайн, за якого запізнюється не більше 5 % подій (формула 2.26).':
        'Chapter 1 metrics that are not among the 18 main ones. “Min. deadline for P ≤ 5 %” is the smallest deadline at which at most 5 % of events are late (formula 2.26).',
    'Порівняння з теорією':
        'Comparison with theory',
    'Середній час перебування подій МРЧ: імітаційне моделювання і три наближення класичної теорії масового обслуговування (розділ 2, п. 2.5.2). Точної формули для моделі з трьома чергами немає, тому відхилення очікувані: наближення 1 занижує, наближення 3 завищує.':
        'Mean sojourn time of SRT events: simulation and three approximations of classical queueing theory (chapter 2, section 2.5.2). There is no exact formula for the three-queue model, so deviations are expected: approximation 1 underestimates, approximation 3 overestimates.',
    'Наближення 1: Поллачек–Хінчин, без пріоритетів':
        'Approximation 1: Pollaczek–Khinchine, no priorities',
    'Наближення 2: Кобгем, відносний пріоритет':
        'Approximation 2: Cobham, non-preemptive priority',
    'Наближення 3: абсолютний пріоритет з дообслуговуванням':
        'Approximation 3: preemptive-resume priority',
    'Мін. дедлайн для P ≤ 5 % (95-й перцентиль), мс':
        'Min. deadline for P ≤ 5 % (95th percentile), ms',
    'Мін. дедлайн для P ≤ 1 % (99-й перцентиль), мс':
        'Min. deadline for P ≤ 1 % (99th percentile), ms',
    'Макс. час реакції (WCRT), мс':
        'Max. response time (WCRT), ms',
    'Джиттер часу реакції (макс. − мін.), мс':
        'Response time jitter (max − min), ms',
    'Найбільше пропусків дедлайну поспіль':
        'Most consecutive deadline misses',
    'Коефіцієнт готовності (без очікування)':
        'Readiness factor (no waiting)',
    '(m, n): найбільше пропусків в одному вікні з n подій':
        '(m, n): most misses in one window of n events',
    '(m, n): частка вікон, де обмеження порушено':
        '(m, n): share of windows violating the constraint',
    '(m, n): частка прогонів без жодного порушення':
        '(m, n): share of runs with no violation',
    'Завантаження процесора U':
        'Processor utilization U',
    'Запас потужності H = 1 − U':
        'Headroom H = 1 − U',
    'Накладні витрати ОС U_os':
        'OS overhead U_os',
    'Ефективність планування η = (U − U_os) / U':
        'Scheduling efficiency η = (U − U_os) / U',
    'Макс. довжина черги Q0 (ISR)':
        'Max. queue length Q0 (ISR)',
    'Макс. довжина черги QH (ЖРЧ)':
        'Max. queue length QH (HRT)',
    'Макс. довжина черги QS (МРЧ)':
        'Max. queue length QS (SRT)',
    'МРЧ: мін. дедлайн для P ≤ 5 %, мс':
        'SRT: min. deadline for P ≤ 5 %, ms',
    'МРЧ: мін. дедлайн для P ≤ 1 %, мс':
        'SRT: min. deadline for P ≤ 1 %, ms',
    'ЖРЧ: мін. дедлайн для P ≤ 1 %, мс':
        'HRT: min. deadline for P ≤ 1 %, ms',
    'МРЧ: джиттер часу реакції, мс':
        'SRT: response time jitter, ms',
    'ЖРЧ: джиттер часу реакції, мс':
        'HRT: response time jitter, ms',
    'МРЧ: найбільше пропусків дедлайну поспіль':
        'SRT: most consecutive deadline misses',
    'ЖРЧ: найбільше пропусків дедлайну поспіль':
        'HRT: most consecutive deadline misses',
    'МРЧ: коефіцієнт готовності':
        'SRT: readiness factor',
    'ЖРЧ: коефіцієнт готовності':
        'HRT: readiness factor',
    'Імітаційне моделювання, мс':
        'Simulation, ms',
    'Додаткові_показники':
        'Additional_metrics',
    'Порівняння_з_теорією':
        'Theory_comparison',
    'Системні_показники':
        'System_metrics',
    'відхилення, %':
        'deviation, %',
    'с':
        's',
}
