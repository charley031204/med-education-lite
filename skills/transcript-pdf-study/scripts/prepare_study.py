"""Extract PT2 once; retain its exact content for the editor and fidelity review."""
import argparse
import re
from pathlib import Path


def extract_pt2(text):
    lines = text.splitlines(keepends=True)
    pattern = re.compile(r'^\s*#{1,6}\s+(?:PART\s*2\b|PT\.?\s*2\b|2부\b)', re.I)
    matches = [i for i, line in enumerate(lines) if pattern.search(line)]
    if len(matches) != 1:
        raise ValueError(f'PT2 heading must be unique; found {len(matches)}')
    body = ''.join(lines[matches[0] + 1:])
    if not body.strip():
        raise ValueError('PT2 is empty')
    if re.search(r'^\s*#{1,6}\s+(?:PART\s*[13-9]\b|PT\.?\s*[13-9]\b|[13-9]부\b)', body, re.I | re.M):
        raise ValueError('Another PART follows PT2; confirm the source boundary')
    return body


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    if args.out.exists():
        parser.error('Output exists; choose a fresh work path')
    body = extract_pt2(args.input.read_text(encoding='utf-8-sig'))
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(body, encoding='utf-8')
    print(f'PT2 extracted: {len(body)} characters; {args.out}')


if __name__ == '__main__':
    main()
