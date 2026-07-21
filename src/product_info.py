from __future__ import annotations

from dataclasses import dataclass
from html import escape
from pathlib import Path
import re


@dataclass(frozen=True)
class ProductFeature:
    title: str
    description: str
    mode: str | None = None


# Единый каталог пользовательских возможностей. Новые крупные блоки нужно
# добавлять сюда одновременно с реализацией: раздел «О программе» строится
# автоматически из этого списка и больше не хранит отдельную ручную копию.
PRODUCT_FEATURES: tuple[ProductFeature, ...] = (
    ProductFeature(
        title="Обычная аналитика продаж",
        mode="Обычный отчет",
        description=(
            "Единый расширенный отчет с уровнем «Проба»: оперативная сводка, показатели "
            "сети, магазинов, камней и номенклатурных групп. Возвраты, отмены и CHAIN "
            "исключаются; денежные показатели отображаются в USD."
        ),
    ),
    ProductFeature(
        title="Сравнение периодов",
        mode="Сравнение периодов",
        description=(
            "Сопоставление двух периодов по сети, магазинам, сегментам, камням и "
            "номенклатурным группам. Глобальные галочки «Серебро / Золото и платина / "
            "Другое» перестраивают все KPI, таблицы и диаграммы; отдельный блок сравнивает "
            "пробы в штуках и выручке с детализацией до камня и группы изделия."
        ),
    ),
    ProductFeature(
        title="Складская аналитика Baserow",
        mode="Сувениры и касты на складе",
        description=(
            "Read-only аналитика сувениров и комплектующих: остатки, фотографии, "
            "категории, материалы, камни, минимальные остатки, проблемные позиции, "
            "движение товара и реестр поставок. Все складские блоки раскрываются одной "
            "страницей, а боковая и мобильная навигация прокручивают к нужному разделу."
        ),
    ),
    ProductFeature(
        title="Заказ Sonu",
        mode="Заказ Sonu",
        description=(
            "Компактный управленческий отчет Sonu из пяти групп: серьги, кольца, подвески, "
            "браслеты не полный круг и браслеты полный круг. Для каждой группы показаны "
            "камни, проданные SKU и изделия, Total продаж, ширина SKU и общий остаток сети, "
            "две диаграммы и рекомендации к заказу без горизонтов в днях. Спорные браслеты "
            "разбираются вручную по крупной фотографии, а сохранённое решение имеет приоритет "
            "над автоматическим правилом 50/50."
        ),
    ),
    ProductFeature(
        title="Заказ поставщику",
        mode="Заказ поставщику",
        description=(
            "Пятый рабочий модуль формирует отдельные заказы по камням и жемчугу. "
            "Он читает ручные комплекты, распределяет их по фактическим продажам, "
            "вычитает остатки 63 и 20, восстанавливает NTR2, отделяет положительный ТВП, "
            "сохраняет черновик, распределяет размеры колец и выгружает поставщику Excel с фотографиями."
        ),
    ),
    ProductFeature(
        title="Единый курс VND/USD",
        description=(
            "Один редактируемый курс применяется к обычному отчету, сравнению "
            "периодов, складской аналитике и заказу Sonu. Базовое значение — "
            "26 300 VND за 1 USD. На телефоне поле курса переносится под подпись и не "
            "перекрывает текст."
        ),
    ),
    ProductFeature(
        title="Интерфейс и стабильность",
        description=(
            "Адаптивная работа на ПК, iPad и смартфонах, единая боковая панель всех модулей "
            "до и после загрузки данных, черно-золотые кнопки с общими состояниями, сортируемые "
            "таблицы, заблокированные от случайного изменения диаграммы, сессионный кэш и защита "
            "одновременного разбора тяжелых Excel-файлов."
        ),
    ),
)

REPORT_MODES: tuple[str, ...] = tuple(
    feature.mode for feature in PRODUCT_FEATURES if feature.mode is not None
)


def feature_cards_html() -> str:
    return "".join(
        '<div class="about-card">'
        f'<h4>{escape(feature.title)}</h4>'
        f'<p>{escape(feature.description)}</p>'
        '</div>'
        for feature in PRODUCT_FEATURES
    )


def release_history_html(changelog_path: Path) -> str:
    """Build release history directly from CHANGELOG.md.

    This removes the second manually maintained version list from the app:
    every new changelog section automatically appears in «О программе».
    """
    try:
        text = changelog_path.read_text(encoding="utf-8")
    except OSError:
        return '<div class="about-step">История обновлений недоступна.</div>'

    sections = re.split(r"(?m)^##\s+", text)
    cards: list[str] = []
    for section in sections[1:]:
        lines = [line.strip() for line in section.strip().splitlines()]
        if not lines:
            continue
        heading = lines[0]
        bullets = [
            re.sub(r"^-\s*", "", line)
            for line in lines[1:]
            if line.startswith("-")
        ]
        if not bullets:
            continue
        summary = " ".join(bullets)
        cards.append(
            '<div class="about-step">'
            f'<b>Analitika Web {escape(heading)}</b><br>'
            f'{escape(summary)}'
            '</div>'
        )
    return "".join(cards)
