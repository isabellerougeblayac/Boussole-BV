import json
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

NS = {
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

AXES = {
    "01_Attenuation": {"id": "01", "key": "ATTENUATION", "label": "Atténuation du changement climatique"},
    "02_Adaptation": {"id": "02", "key": "ADAPTATION", "label": "Adaptation au changement climatique et prévention des risques naturels"},
    "03_Eau": {"id": "03", "key": "EAU", "label": "Gestion des ressources en eau"},
    "04_Eco circ": {"id": "04", "key": "ECO_CIRC", "label": "Économie circulaire, déchets et risques technologiques"},
    "05_Pollution": {"id": "05", "key": "POLLUTION", "label": "Prévention et contrôle des pollutions de l'air et des sols"},
    "06_Biodiversité": {"id": "06", "key": "BIODIVERSITE", "label": "Préservation de la biodiversité et protection des espaces naturels"},
}

RATING_ORDER = {
    "très favorable": 0,
    "favorable sous conditions": 1,
    "neutre": 2,
    "défavorable": 3,
    "non cotée": 4,
    "hors périmètre": 5,
}


def clean_text(value):
    return re.sub(r"[ \t\r\f\v]+", " ", (value or "").replace("\xa0", " ")).strip()


def normalized_key(value):
    return clean_text(value).casefold()


def readable_label(labels):
    counts = Counter(labels)
    return sorted(counts, key=lambda label: (-counts[label], label.isupper(), label.casefold()))[0]


def domain_sort_key(value):
    key = normalized_key(value)
    return (key == "hors périmètre", key)


def rating_sort_key(value):
    return RATING_ORDER.get(normalized_key(value), len(RATING_ORDER))


def normalize_criteria(value):
    text = clean_text(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = []
    for raw_line in text.split("\n"):
        line = clean_text(raw_line)
        if not line:
            continue
        line = re.sub(r"^[•◦▪▫‣⁃]+\s*", "- ", line)
        line = re.sub(r"^[–—]\s*", "- ", line)
        line = re.sub(r"^[-]\s+", "- ", line)
        line = re.sub(r"\s+[–—]\s+", " - ", line)
        lines.append(line)
    return "\n".join(lines)


def cell_text(cell):
    parts = []
    for paragraph in cell.findall(".//text:p", NS):
        text = "".join(paragraph.itertext()).strip()
        if text:
            parts.append(text)
    return clean_text("\n".join(parts))


def row_values(row, max_cols=5):
    values = []
    for cell in list(row):
        if not cell.tag.endswith("table-cell") and not cell.tag.endswith("covered-table-cell"):
            continue
        repeat = int(cell.attrib.get(f"{{{NS['table']}}}number-columns-repeated", "1"))
        value = cell_text(cell)
        for _ in range(repeat):
            values.append(value)
            if len(values) >= max_cols:
                return values
    return values


def main():
    source = Path("20260727_6_Axes_integration.ods")
    output = Path("src/data/rules.json")

    with zipfile.ZipFile(source) as archive:
        root = ET.fromstring(archive.read("content.xml"))

    raw_rules = []
    domain_labels = defaultdict(list)
    type_labels = defaultdict(lambda: defaultdict(list))

    for table in root.findall(".//table:table", NS):
        sheet_name = table.attrib.get(f"{{{NS['table']}}}name")
        axis = AXES.get(sheet_name)
        if not axis:
            continue

        for index, row in enumerate(table.findall("table:table-row", NS)):
            if index == 0:
                continue
            values = row_values(row, 5)
            values += [""] * (5 - len(values))
            domain, expense_type, criteria, rating, criterion_id = [clean_text(value) for value in values[:5]]
            if not domain or not expense_type or not criteria or not rating or not criterion_id:
                continue

            domain_key = normalized_key(domain)
            type_key = normalized_key(expense_type)
            domain_labels[domain_key].append(domain)
            type_labels[domain_key][type_key].append(expense_type)
            raw_rules.append(
                {
                    "axis": axis,
                    "domainKey": domain_key,
                    "expenseTypeKey": type_key,
                    "criteria": normalize_criteria(criteria),
                    "rating": rating,
                    "criterionId": criterion_id,
                    "sourceIndex": len(raw_rules),
                }
            )

    canonical_domains = {key: readable_label(labels) for key, labels in domain_labels.items()}
    canonical_types = {
        domain_key: {type_key: readable_label(labels) for type_key, labels in values.items()}
        for domain_key, values in type_labels.items()
    }

    rules = []
    for item in raw_rules:
        axis = item["axis"]
        rules.append(
            {
                "axisId": axis["id"],
                "axisKey": axis["key"],
                "axisLabel": axis["label"],
                "domain": canonical_domains[item["domainKey"]],
                "domainKey": item["domainKey"],
                "expenseType": canonical_types[item["domainKey"]][item["expenseTypeKey"]],
                "expenseTypeKey": item["expenseTypeKey"],
                "criteria": item["criteria"],
                "rating": item["rating"],
                "criterionId": item["criterionId"],
            }
        )

    rules.sort(
        key=lambda rule: (
            domain_sort_key(rule["domain"]),
            normalized_key(rule["expenseType"]),
            rule["axisId"],
            rating_sort_key(rule["rating"]),
            rule["criterionId"],
        )
    )

    sorted_domain_keys = sorted(canonical_domains, key=lambda key: domain_sort_key(canonical_domains[key]))
    payload = {
        "axes": list(AXES.values()),
        "domains": [canonical_domains[key] for key in sorted_domain_keys],
        "typesByDomain": {
            canonical_domains[domain_key]: [
                canonical_types[domain_key][type_key]
                for type_key in sorted(canonical_types[domain_key], key=lambda key: normalized_key(canonical_types[domain_key][key]))
            ]
            for domain_key in sorted_domain_keys
        },
        "rules": rules,
    }

    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {output} with {len(rules)} rules")


if __name__ == "__main__":
    main()
