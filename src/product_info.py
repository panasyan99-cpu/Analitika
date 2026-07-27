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
        title="Анализ продаж",
        mode="Обычный отчет",
        description=(
            "Короткая сводка сети, лидер розничной сети, магазины, камни и группы, "
            "поставщики и интерактивные фильтры. Денежные показатели отображаются в USD."
        ),
    ),
    ProductFeature(
        title="Сравнение периодов",
        mode="Сравнение периодов",
        description=(
            "Итоги двух периодов, продажи в день, драйверы роста и снижения, новые и "
            "исчезнувшие группы, магазины, камни, металлы, пробы и поставщики."
        ),
    ),
    ProductFeature(
        title="Склад Baserow",
        mode="Сувениры и касты на складе",
        description=(
            "Остатки сувениров и комплектующих, фотографии, минимальные остатки, "
            "движение товара, поставки и позиции, требующие внимания."
        ),
    ),
    ProductFeature(
        title="Заказ Sonu",
        mode="Заказ Sonu",
        description=(
            "Продажи, сетевые остатки и рекомендации по пяти товарным группам Sonu, "
            "включая ручной разбор спорных моделей браслетов."
        ),
    ),
    ProductFeature(
        title="Заказ поставщику",
        mode="Заказ поставщику",
        description=(
            "Заказы по камням и жемчугу, рекомендации, ручные количества, размеры колец, "
            "замена замков, Limited Order, черновики, история и готовый Excel."
        ),
    ),
    ProductFeature(
        title="Документация",
        description=(
            "Краткое описание возможностей, полное руководство пользователя, история "
            "обновлений и технический README для поддержки проекта."
        ),
    ),
)

REPORT_MODES: tuple[str, ...] = (
    "Обычный отчет",
    "Сравнение периодов",
    "Сувениры и касты на складе",
    "Заказ Sonu",
    "Заказ поставщику",
    "О программе",
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
