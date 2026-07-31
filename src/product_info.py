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
            "Сводный обзор результатов сети: ключевые показатели, лидер розничной сети, "
            "структура ассортимента, магазины, категории и поставщики."
        ),
    ),
    ProductFeature(
        title="Сравнение периодов",
        mode="Сравнение периодов",
        description=(
            "Сопоставление результатов двух периодов с объяснением динамики, драйверов "
            "роста и снижения, изменений по магазинам, ассортименту и поставщикам."
        ),
    ),
    ProductFeature(
        title="Склад Baserow",
        mode="Сувениры и касты на складе",
        description=(
            "Контроль остатков, приёмки, передачи в бухгалтерию, движения товара, "
            "поставок и позиций, требующих управленческого внимания."
        ),
    ),
    ProductFeature(
        title="Заказ Sonu",
        mode="Заказ Sonu",
        description=(
            "Анализ ассортимента Sonu, продаж и сетевых остатков с рекомендациями по "
            "пополнению и ручным разбором спорных моделей браслетов."
        ),
    ),
    ProductFeature(
        title="Заказ поставщику",
        mode="Заказ поставщику",
        description=(
            "Расчёт потребности по камням и жемчугу: рекомендации, ручные количества, "
            "размеры колец, замки, Limited Order, история и готовый Excel."
        ),
    ),
    ProductFeature(
        title="Управленческий отчет",
        mode="Управленческий отчет",
        description=(
            "Сравнение двух одинаковых периодов по общим KPI, показателям в день, "
            "магазинам, продавцам, поставщикам, категориям, камням, пробам, SKU и возвратам."
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
    "Управленческий отчет",
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
