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
            "Оперативная сводка для руководителя, показатели сети, детализация "
            "магазинов, интерактивные срезы и аналитика поставщиков. Возвраты, "
            "отмены и CHAIN исключаются; денежные показатели отображаются в USD."
        ),
    ),
    ProductFeature(
        title="Сравнение периодов",
        mode="Сравнение периодов",
        description=(
            "Сопоставление двух периодов по сети, магазинам, сегментам, камням, "
            "номенклатурным группам и поставщикам. Система проверяет периоды, "
            "ставит более ранний слева и показывает абсолютные и процентные изменения."
        ),
    ),
    ProductFeature(
        title="Складская аналитика Baserow",
        mode="Сувениры и касты на складе",
        description=(
            "Read-only аналитика сувениров и комплектующих: остатки, фотографии, "
            "категории, материалы, камни, минимальные остатки, проблемные позиции, "
            "движение товара и реестр поставок."
        ),
    ),
    ProductFeature(
        title="Заказ Sonu",
        mode="Заказ Sonu",
        description=(
            "Продажи по магазинам, средние продажи и взвешенная средняя цена, "
            "разбор браслетов и камней, прогноз продаж, матрица заказа и итоговые "
            "рекомендации. Возвраты не учитываются; все суммы выводятся в USD."
        ),
    ),
    ProductFeature(
        title="Единый курс VND/USD",
        description=(
            "Один редактируемый курс применяется к обычному отчету, сравнению "
            "периодов, складской аналитике и заказу Sonu. Базовое значение — "
            "26 300 VND за 1 USD."
        ),
    ),
    ProductFeature(
        title="Интерфейс и стабильность",
        description=(
            "Адаптивная работа на ПК, iPad и смартфонах, фирменная навигация, "
            "сортируемые таблицы, заблокированные от случайного изменения диаграммы, "
            "сессионный кэш и защита одновременного разбора тяжелых Excel-файлов."
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
