#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Corrige e valida a hierarquia das entidades da Base 4711.

Modelo esperado:
G4F
├── 1.1 - ...
│   ├── 1.1.1 - entidade operacional unificada
│   └── 1.1.3 - Atividade de Liderança
├── 1.2 - ...
│   ├── 1.2.1 - entidade operacional unificada
│   └── 1.2.3 - Atividade de Liderança
...
└── 6.1 - ...
    ├── 6.1.1 - entidade operacional unificada
    └── 6.1.3 - Atividade de Liderança

Regras:
- G4F é a entidade raiz (ID 0);
- cada entidade de serviço X.Y fica diretamente abaixo de G4F;
- cada entidade operacional X.Y.1 fica diretamente abaixo de X.Y;
- cada entidade de liderança X.Y.3 fica diretamente abaixo de X.Y;
- entidades X.Y.2 já migradas/lixeira não são recriadas;
- IDs e nomes existentes são preservados, alterando somente entities_id quando necessário.

Uso:
  source /root/.glpi.env
  python3 scripts/fix_entity_hierarchy.py --dry-run
  python3 scripts/fix_entity_hierarchy.py --apply
"""

import argparse
import re
import sys
from typing import Any, Dict, List, Optional

from import_base4711 import (
    GLPI,
    ROOT_ENTITY_ID,
    clean_text,
)


SERVICE_CODES = [
    "1.1", "1.2", "1.3",
    "2.1", "2.2", "2.3",
    "3.1", "3.2", "3.3",
    "4.1", "4.2", "4.3",
    "5.1", "5.2", "5.3",
    "6.1",
]


def entity_code(name: str) -> str:
    match = re.match(
        r"^\s*(\d+\.\d+(?:\.\d+)?)\s*[-–—]\s*",
        clean_text(name),
    )
    return match.group(1) if match else ""


def parent_id(entity: Dict[str, Any]) -> int:
    return int(entity.get("entities_id") or 0)


def find_unique(
    entities: List[Dict[str, Any]],
    code: str,
    *,
    required: bool = True,
) -> Optional[Dict[str, Any]]:
    matches = [
        entity
        for entity in entities
        if entity_code(clean_text(entity.get("name"))) == code
    ]

    if len(matches) == 1:
        return matches[0]

    if len(matches) == 0 and not required:
        return None

    if len(matches) == 0:
        raise RuntimeError(
            f"Entidade '{code} - ...' não localizada."
        )

    raise RuntimeError(
        f"Mais de uma entidade localizada para '{code}': "
        + ", ".join(
            f"ID={x.get('id')} '{x.get('name')}'"
            for x in matches
        )
    )


def build_plan(
    entities: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    plan: List[Dict[str, Any]] = []

    for service_code in SERVICE_CODES:
        service = find_unique(
            entities,
            service_code,
            required=True,
        )
        service_id = int(service["id"])

        operational = find_unique(
            entities,
            f"{service_code}.1",
            required=True,
        )
        leadership = find_unique(
            entities,
            f"{service_code}.3",
            required=True,
        )

        plan.append(
            {
                "service_code": service_code,
                "service": service,
                "operational": operational,
                "leadership": leadership,
                "service_expected_parent": ROOT_ENTITY_ID,
                "operational_expected_parent": service_id,
                "leadership_expected_parent": service_id,
            }
        )

    return plan


def print_plan(plan: List[Dict[str, Any]]) -> int:
    changes = 0

    print()
    print("=" * 112)
    print("VALIDAÇÃO DA HIERARQUIA DE ENTIDADES")
    print("=" * 112)

    for item in plan:
        service = item["service"]
        operational = item["operational"]
        leadership = item["leadership"]

        service_current = parent_id(service)
        operational_current = parent_id(operational)
        leadership_current = parent_id(leadership)

        service_expected = int(item["service_expected_parent"])
        operational_expected = int(item["operational_expected_parent"])
        leadership_expected = int(item["leadership_expected_parent"])

        service_ok = service_current == service_expected
        operational_ok = operational_current == operational_expected
        leadership_ok = leadership_current == leadership_expected

        changes += int(not service_ok)
        changes += int(not operational_ok)
        changes += int(not leadership_ok)

        print()
        print(
            f"{item['service_code']} | "
            f"ID={service.get('id')} {service.get('name')}"
        )
        print(
            "  Pai do serviço..............: "
            f"atual={service_current} esperado={service_expected} "
            f"{'OK' if service_ok else 'CORRIGIR'}"
        )
        print(
            f"  Filho operacional .1........: "
            f"ID={operational.get('id')} "
            f"pai_atual={operational_current} "
            f"pai_esperado={operational_expected} "
            f"{'OK' if operational_ok else 'CORRIGIR'}"
        )
        print(
            f"  Filho liderança .3..........: "
            f"ID={leadership.get('id')} "
            f"pai_atual={leadership_current} "
            f"pai_esperado={leadership_expected} "
            f"{'OK' if leadership_ok else 'CORRIGIR'}"
        )

    print()
    print("-" * 112)
    print(f"Serviços validados....................: {len(plan)}")
    print(f"Relacionamentos a corrigir............: {changes}")
    print(
        "Estrutura esperada....................: "
        "G4F > X.Y > X.Y.1 / X.Y.3"
    )
    if changes == 0:
        print("Status.................................: HIERARQUIA JÁ CORRETA")
    print("=" * 112)

    return changes


def apply_plan(
    glpi: GLPI,
    plan: List[Dict[str, Any]],
) -> int:
    changed = 0

    for item in plan:
        service = item["service"]
        operational = item["operational"]
        leadership = item["leadership"]

        targets = [
            (
                service,
                int(item["service_expected_parent"]),
                "serviço",
            ),
            (
                operational,
                int(item["operational_expected_parent"]),
                "operacional .1",
            ),
            (
                leadership,
                int(item["leadership_expected_parent"]),
                "liderança .3",
            ),
        ]

        for entity, expected_parent, label in targets:
            current_parent = parent_id(entity)

            if current_parent == expected_parent:
                continue

            entity_id = int(entity["id"])
            glpi.update(
                "Entity",
                entity_id,
                {
                    "entities_id": expected_parent,
                },
            )
            entity["entities_id"] = expected_parent
            changed += 1

            print(
                f"CORRIGIDO {label}: Entity ID={entity_id} "
                f"pai {current_parent} -> {expected_parent}"
            )

    print()
    print("=" * 112)
    print("CORREÇÃO DE HIERARQUIA CONCLUÍDA")
    print("=" * 112)
    print(f"Relacionamentos corrigidos............: {changed}")
    print("=" * 112)

    return changed


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Valida/corrige a hierarquia G4F > serviço > operacional/liderança."
        )
    )

    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--dry-run", action="store_true")
    mode.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    glpi = GLPI()

    try:
        glpi.init()

        entities = glpi.get_all("Entity")
        plan = build_plan(entities)
        changes = print_plan(plan)

        if args.dry_run:
            print()
            print("DRY-RUN concluído. Nenhuma alteração foi realizada.")
            return 0

        if changes == 0:
            print()
            print("Nenhuma alteração necessária.")
            return 0

        apply_plan(glpi, plan)
        return 0

    finally:
        glpi.close()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nExecução interrompida.", file=sys.stderr)
        sys.exit(130)
    except Exception as exc:
        print(f"\nERRO: {exc}", file=sys.stderr)
        sys.exit(1)
