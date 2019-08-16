import json


def save_as_json(path, lines):
    with open(path, 'w', encoding='utf8') as outfile:
        outfile.write(json.dumps(lines, ensure_ascii=False))


def main():
    f = open('prof.txt', 'r')
    lines = f.readlines()
    lines = list(map(lambda s: s.replace('\n', ''), lines))
    save_as_json('prof.json', lines)


if __name__ == "__main__":
    main()
