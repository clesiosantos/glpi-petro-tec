#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Migra a árvore de entidades da Base 4711 para o modelo operacional unificado.

Objetivo:
- para cada serviço (1.1, 1.2, ..., 6.1), unifica as subentidades .1 e .2;
- reutiliza a entidade .1 como destino, preservando seu ID;
- renomeia .1 com os dois escopos combinados;
- migra categorias ITIL associadas à antiga .2 para a nova .1;
- migra autorizações do perfil "Posto de Trabalho" do pai e da antiga .2
  para a nova .1, com recursividade = Sim;
- ajusta a entidade padrão dos usuários Posto de Trabalho;
- envia a antiga entidade .2 para a lixeira;
- mantém a subentidade .3 (Liderança) inalterada.

A árvore de categorias NÃO é fundida neste script. As categorias continuam com
seus nomes/códigos atuais, mas passam a apontar para a entidade operacional
unificada .1.

Uso:
  source /root/.glpi.env
  python3 scripts/migrate_unify_operational_entities.py --dry-run
  python3 scripts/migrate_unify_operational_entities.py --apply
"""

import argparse
import re
import sys
from typing import Any, Dict, List, Optional, Tuple

from import_base4711 import (
    API_URL,
    GLPI,
    ROOT_ENTITY_ID,
    clean_text,
    get_profiles_by_name,
    normalize_compare,
)


SERVICE_CODES = [
    "1.1", "1.2", "1.3",
    "2.1", "2.2", "2.3",
    "3.1", "3.2", "3.3",
    "4.1", "4.2", "4.3",
    "5.1", "5.2", "5.3",
    "6.1",
]

POSTO_PROFILE = "Posto de Trabalho"


def entity_code(name: str) -> str:
    match = re.match(r"^\s*(\d+\.\d+(?:\.\d+)?)\s*[-–—]\s*", clean_text(name))
    return match.group(1) if match else ""


def entity_suffix(name: str) -> str:
    text = clean_text(name)
    match = re.match(r"^\s*\d+\.\d+(?:\.\d+)?\s*[-–—]\s*(.+)$", text)
    return clean_text(match.group(1)) if match else text


def entity_parent_id(entity: Dict[str, Any]) -> int:
    return int(entity.get("entities_id") or 0)


def find_single_entity(
    entities: List[Dict[str, Any]],
    code: str,
    *,
    parent_id: Optional[int] = None,
    required: bool = True,
) -> Optional[Dict[str, Any]]:
    code_matches = [
        e
        for e in entities
        if entity_code(clean_text(e.get("name"))) == code
    ]

    if parent_id is not None:
        matches = [
            e
            for e in code_matches
            if entity_parent_id(e) == int(parent_id)
        ]
    else:
        matches = code_matches

    if len(matches) == 1:
        return matches[0]

    # Pós-migração: em alguns retornos da API o parent pode não vir
    # materializado como antes. Se o código for único globalmente,
    # usamos o item único e exibimos o parent real no relatório.
    if (
        parent_id is not None
        and len(matches) == 0
        and len(code_matches) == 1
    ):
        return code_matches[0]

    if len(matches) == 0 and not required:
        return None

    if len(matches) == 0:
        raise RuntimeError(
            f"Entidade '{code} - ...' não localizada"
            + (
                f" sob parent ID={parent_id}."
                if parent_id is not None
                else "."
            )
        )

    raise RuntimeError(
        f"Mais de uma entidade encontrada para código {code}: "
        + ", ".join(
            f"ID={e.get('id')} '{e.get('name')}'"
            for e in matches
        )
    )


def build_plan(
    entities: List[Dict[str, Any]],
    categories: List[Dict[str, Any]],
    profile_users: List[Dict[str, Any]],
    users: List[Dict[str, Any]],
    posto_profile_id: int,
) -> List[Dict[str, Any]]:
    plan: List[Dict[str, Any]] = []

    for service_code in SERVICE_CODES:
        parent = find_single_entity(
            entities,
            service_code,
            parent_id=ROOT_ENTITY_ID,
            required=True,
        )
        parent_id = int(parent["id"])

        target_code = f"{service_code}.1"
        source_code = f"{service_code}.2"
        leadership_code = f"{service_code}.3"

        target = find_single_entity(
            entities,
            target_code,
            parent_id=parent_id,
            required=True,
        )
        source = find_single_entity(
            entities,
            source_code,
            parent_id=parent_id,
            required=False,
        )
        leadership = find_single_entity(
            entities,
            leadership_code,
            parent_id=parent_id,
            required=False,
        )

        target_id = int(target["id"])
        source_id = int(source["id"]) if source else 0

        if source:
            source_children = [
                e
                for e in entities
                if entity_parent_id(e) == source_id
            ]
            if source_children:
                raise RuntimeError(
                    f"A entidade {source_code} ID={source_id} possui "
                    f"{len(source_children)} subentidade(s). "
                    "Migração interrompida para evitar perda de hierarquia."
                )

            suffix1 = entity_suffix(clean_text(target.get("name")))
            suffix2 = entity_suffix(clean_text(source.get("name")))

            if normalize_compare(suffix2) not in normalize_compare(suffix1):
                unified_name = (
                    f"{target_code} - {suffix1} / {suffix2}"
                )
            else:
                unified_name = clean_text(target.get("name"))
        else:
            unified_name = clean_text(target.get("name"))

        category_rows = [
            x
            for x in categories
            if source_id > 0
            and int(x.get("entities_id") or 0) == source_id
        ]

        profile_rows = [
            x
            for x in profile_users
            if int(x.get("profiles_id") or 0) == posto_profile_id
            and int(x.get("entities_id") or 0)
            in ({parent_id, source_id} if source_id else {parent_id})
        ]

        user_rows = [
            x
            for x in users
            if int(x.get("profiles_id") or 0) == posto_profile_id
            and int(x.get("entities_id") or 0)
            in ({parent_id, source_id} if source_id else {parent_id})
        ]

        plan.append(
            {
                "service_code": service_code,
                "parent": parent,
                "target": target,
                "source": source,
                "leadership": leadership,
                "unified_name": unified_name,
                "categories": category_rows,
                "profile_users": profile_rows,
                "users": user_rows,
            }
        )

    return plan


def print_plan(plan: List[Dict[str, Any]]) -> None:
    total_categories = 0
    total_auth = 0
    total_users = 0
    total_sources = 0

    print()
    print("=" * 112)
    print("PLANO DE UNIFICAÇÃO DAS SUBENTIDADES OPERACIONAIS")
    print("=" * 112)

    for item in plan:
        parent = item["parent"]
        target = item["target"]
        source = item["source"]
        leadership = item["leadership"]

        categories = len(item["categories"])
        auths = len(item["profile_users"])
        users = len(item["users"])

        total_categories += categories
        total_auth += auths
        total_users += users

        if source:
            total_sources += 1
            source_text = (
                f"ID={source.get('id')} {source.get('name')}"
            )
        else:
            source_text = "JÁ MIGRADA / .2 não ativa"

        print()
        print(
            f"{item['service_code']} | "
            f"Pai ID={parent.get('id')} {parent.get('name')}"
        )
        print(
            f"  DESTINO .1 : ID={target.get('id')} "
            f"{target.get('name')}"
        )
        print(f"  ORIGEM  .2 : {source_text}")
        print(
            f"  NOVO NOME  : {item['unified_name']}"
        )
        print(
            f"  Categorias .2 a migrar.............: {categories}"
        )
        print(
            f"  Autorizações Posto a migrar........: {auths}"
        )
        print(
            f"  Usuários com entidade padrão ajustar: {users}"
        )
        if leadership:
            print(
                f"  Liderança preservada.................: "
                f"ID={leadership.get('id')} {leadership.get('name')}"
            )

    print()
    print("-" * 112)
    print(f"Serviços processados..................: {len(plan)}")
    print(f"Entidades .2 a enviar para lixeira....: {total_sources}")
    print(f"Categorias a migrar...................: {total_categories}")
    print(f"Autorizações a migrar.................: {total_auth}")
    print(f"Usuários padrão a ajustar.............: {total_users}")
    if (
        total_sources == 0
        and total_categories == 0
        and total_auth == 0
        and total_users == 0
    ):
        print("Status.................................: MIGRAÇÃO JÁ CONCLUÍDA")
    print("=" * 112)


def migrate_profile_users(
    glpi: GLPI,
    rows: List[Dict[str, Any]],
    all_profile_users: List[Dict[str, Any]],
    target_id: int,
) -> int:
    changed = 0

    existing_keys = {
        (
            int(x.get("users_id") or 0),
            int(x.get("profiles_id") or 0),
            int(x.get("entities_id") or 0),
            int(x.get("is_recursive") or 0),
        )
        for x in all_profile_users
    }

    for auth in rows:
        auth_id = int(auth["id"])
        user_id = int(auth.get("users_id") or 0)
        profile_id = int(auth.get("profiles_id") or 0)
        dest_key = (user_id, profile_id, target_id, 1)

        if dest_key in existing_keys:
            glpi.purge("Profile_User", auth_id)
            print(
                f"  PURGE Profile_User ID={auth_id}: "
                f"destino já existia user={user_id}"
            )
        else:
            glpi.update(
                "Profile_User",
                auth_id,
                {
                    "entities_id": target_id,
                    "is_recursive": 1,
                    "is_dynamic": 0,
                },
            )
            existing_keys.add(dest_key)
            print(
                f"  MOVE Profile_User ID={auth_id}: "
                f"user={user_id} -> entity={target_id} recursivo=1"
            )

        changed += 1

    return changed


def trash_entity(glpi: GLPI, entity_id: int) -> None:
    r = glpi.http.delete(
        f"{API_URL}/Entity/{entity_id}",
        params={"force_purge": "false"},
        timeout=60,
    )
    if r.status_code not in (200, 201, 204):
        raise RuntimeError(
            f"DELETE Entity/{entity_id}: "
            f"HTTP {r.status_code}: {r.text}"
        )


def apply_plan(
    glpi: GLPI,
    plan: List[Dict[str, Any]],
    categories: List[Dict[str, Any]],
    profile_users: List[Dict[str, Any]],
) -> None:
    renamed = 0
    moved_categories = 0
    moved_auth = 0
    moved_users = 0
    trashed_entities = 0

    for item in plan:
        target = item["target"]
        source = item["source"]
        target_id = int(target["id"])
        current_name = clean_text(target.get("name"))
        unified_name = item["unified_name"]

        print()
        print(
            f"=== {item['service_code']} -> Entity ID {target_id} ==="
        )

        if current_name != unified_name:
            glpi.update(
                "Entity",
                target_id,
                {"name": unified_name},
            )
            target["name"] = unified_name
            renamed += 1
            print(
                f"  RENOMEADA entidade .1: '{unified_name}'"
            )

        if source:
            source_id = int(source["id"])

            for category in item["categories"]:
                category_id = int(category["id"])
                glpi.update(
                    "ITILCategory",
                    category_id,
                    {
                        "entities_id": target_id,
                    },
                )
                category["entities_id"] = target_id
                moved_categories += 1
                print(
                    f"  MOVE ITILCategory ID={category_id}: "
                    f"entity {source_id} -> {target_id}"
                )

            moved_auth += migrate_profile_users(
                glpi,
                item["profile_users"],
                profile_users,
                target_id,
            )

            for user in item["users"]:
                user_id = int(user["id"])
                glpi.update(
                    "User",
                    user_id,
                    {"entities_id": target_id},
                )
                user["entities_id"] = target_id
                moved_users += 1
                print(
                    f"  MOVE User ID={user_id}: "
                    f"entidade padrão -> {target_id}"
                )

            trash_entity(glpi, source_id)
            trashed_entities += 1
            print(
                f"  LIXEIRA Entity ID={source_id}: "
                f"{source.get('name')}"
            )

    print()
    print("=" * 112)
    print("MIGRAÇÃO CONCLUÍDA")
    print("=" * 112)
    print(f"Entidades .1 renomeadas...............: {renamed}")
    print(f"Categorias migradas...................: {moved_categories}")
    print(f"Autorizações migradas.................: {moved_auth}")
    print(f"Usuários padrão ajustados.............: {moved_users}")
    print(f"Entidades .2 enviadas para lixeira....: {trashed_entities}")
    print("=" * 112)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Unifica subentidades .1 e .2 da árvore operacional Base 4711."
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
        categories = glpi.get_all("ITILCategory")
        profiles = glpi.get_all("Profile")
        profile_users = glpi.get_all("Profile_User")
        users = glpi.get_all("User")

        profiles_by_name = get_profiles_by_name(profiles)
        posto_profile_id = profiles_by_name.get(
            normalize_compare(POSTO_PROFILE)
        )
        if not posto_profile_id:
            raise RuntimeError(
                f"Perfil '{POSTO_PROFILE}' não localizado."
            )

        plan = build_plan(
            entities,
            categories,
            profile_users,
            users,
            posto_profile_id,
        )
        print_plan(plan)

        if args.dry_run:
            print()
            print("DRY-RUN concluído. Nenhuma alteração foi realizada.")
            return 0

        apply_plan(
            glpi,
            plan,
            categories,
            profile_users,
        )
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
