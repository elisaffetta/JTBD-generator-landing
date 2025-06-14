# Пример промпта для генерации JTBD

Этот файл содержит пример того, как должен выглядеть промпт для генерации качественных Jobs To Be Done.

## Базовый промпт (используется в скрипте)

```
Ты эксперт по Jobs To Be Done (JTBD) методологии. Твоя задача - проанализировать информацию о продукте/нише/ЦА и сгенерировать 10 реальных, конкретных и практичных JTBD.

Информация о продукте/нише/ЦА: [ИНФОРМАЦИЯ ОТ ПОЛЬЗОВАТЕЛЯ]

Для каждого JTBD предоставь:
1. Формулировку JTBD в формате "Когда я [ситуация], я хочу [мотивация], чтобы [ожидаемый результат]"
2. Краткое обоснование (2-3 предложения) - откуда взялась эта потребность и почему она важна для данной ЦА
3. Концепцию Google Sheets инструмента для решения этой задачи (1-2 предложения)
4. Концепцию веб-приложения/мини-инструмента для решения этой задачи (1-2 предложения, фокус на 1 продукт = 1 JTBD = 1 ЦА)

Формат ответа должен быть JSON:
{
  "jtbds": [
    {
      "jtbd": "текст JTBD",
      "reasoning": "обоснование",
      "sheets_solution": "концепция Google Sheets решения",
      "web_solution": "концепция веб-приложения"
    }
  ]
}

Важно: JTBD должны быть реалистичными, основанными на реальных потребностях пользователей, а не надуманными.
```

## Расширенный промпт (для более детальной генерации)

Если вы хотите получить более детальные результаты, можете заменить промпт в коде на этот:

```
Ты эксперт по Jobs To Be Done (JTBD) методологии и продуктовой аналитике. Твоя задача - глубоко проанализировать информацию о продукте/нише/ЦА и сгенерировать 10 реальных, конкретных и практичных JTBD, основанных на реальных потребностях пользователей.

Информация о продукте/нише/ЦА: [ИНФОРМАЦИЯ ОТ ПОЛЬЗОВАТЕЛЯ]

Методология анализа:
1. Определи ключевые болевые точки целевой аудитории
2. Выяви контексты использования продукта
3. Найди моменты фрустрации в пользовательском пути
4. Определи желаемые результаты пользователей
5. Сформулируй JTBD на основе реальных сценариев использования

Для каждого JTBD предоставь:
1. **Формулировку JTBD** в формате "Когда я [конкретная ситуация], я хочу [четкая мотивация], чтобы [измеримый результат]"
2. **Детальное обоснование** (3-4 предложения):
   - Откуда взялась эта потребность
   - Почему она критична для данной ЦА
   - Какие альтернативные решения существуют сейчас
   - Почему текущие решения неудовлетворительны
3. **Google Sheets решение** (2-3 предложения):
   - Конкретный функционал
   - Ключевые формулы или автоматизации
   - Преимущества такого подхода
4. **Веб-приложение концепция** (2-3 предложения):
   - Фокус на решении одной конкретной задачи
   - Уникальное ценностное предложение
   - Монетизационная модель

Требования к качеству:
- JTBD должны быть основаны на реальных исследованиях пользователей
- Избегай общих формулировок типа "хочу быть эффективнее"
- Фокусируйся на конкретных ситуациях и измеримых результатах
- Учитывай эмоциональные и социальные аспекты потребностей

Формат ответа - строгий JSON:
{
  "jtbds": [
    {
      "jtbd": "детальная формулировка JTBD",
      "reasoning": "подробное обоснование с анализом",
      "sheets_solution": "конкретная концепция Google Sheets инструмента",
      "web_solution": "детальная концепция веб-приложения с УТП"
    }
  ]
}
```

## Как изменить промпт в скрипте

1. Откройте файл `JTBD_Generator.gs` в Apps Script
2. Найдите функцию `generateJTBDWithAPI`
3. Найдите переменную `prompt`
4. Замените содержимое на желаемый промпт
5. Сохраните изменения

## Примеры качественных JTBD

### Хорошие примеры:
- "Когда я планирую недельное меню для семьи, я хочу учесть диетические ограничения каждого члена семьи, чтобы сэкономить время на походы в магазин и избежать пищевых отходов"
- "Когда я веду переговоры с клиентом по телефону, я хочу быстро найти его историю покупок, чтобы предложить персонализированное решение и увеличить вероятность сделки"

### Плохие примеры:
- "Когда я работаю, я хочу быть продуктивнее" (слишком общее)
- "Когда я использую приложение, я хочу удобный интерфейс" (фокус на решении, а не на потребности)

## Советы по оптимизации промпта

1. **Добавьте контекст отрасли** - если работаете в специфической нише
2. **Укажите географию** - для учета культурных особенностей
3. **Определите временные рамки** - краткосрочные vs долгосрочные потребности
4. **Добавьте ограничения** - бюджет, время, ресурсы ЦА
5. **Включите конкурентный анализ** - что уже существует на рынке