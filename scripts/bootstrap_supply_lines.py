from __future__ import annotations

import getpass
import json
from pathlib import Path
import tomllib

from src.warehouse_management.schema import BaserowSchemaManager


def main() -> None:
    secrets_path = Path('.streamlit/secrets.toml')
    if not secrets_path.exists():
        raise SystemExit('Не найден .streamlit/secrets.toml')
    secrets = tomllib.loads(secrets_path.read_text('utf-8'))
    config = secrets['baserow']
    email = str(config.get('email') or input('Email Baserow: ')).strip()
    password = str(config.get('password') or getpass.getpass('Пароль Baserow: '))
    manager = BaserowSchemaManager(str(config['url']), email, password)
    report = manager.ensure_and_migrate(
        database_id=int(config['database_id']),
        souvenirs_table_id=int(config['souvenirs_table_id']),
        components_table_id=int(config['components_table_id']),
        operations_table_id=int(config['operations_table_id']),
        supplies_table_id=int(config['supplies_table_id']),
    )
    Path('warehouse_supply_lines_migration_2.4.0.json').write_text(
        json.dumps(report.to_dict(), ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    print(f'Позиции поставок готовы. Table ID: {report.table_id}')
    print(f'Перенесено строк: {report.migrated_lines}')
    print('Добавьте/обновите в secrets.toml:')
    print(f'supply_lines_table_id = {report.table_id}')


if __name__ == '__main__':
    main()
