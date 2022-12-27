import argparse
from pathlib import Path

from .word_parser import WordParser
from config import QUESTIONNAIRE_FILE

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Parse questionnaire from text files')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-f', '--file', action='store_true', help='parse from file')
    group.add_argument('-d', '--dir', action='store_true',
                       help='parse all files appropriate type in directory '
                            '(example: *.docx)')
    parser.add_argument('-t', '--filetype', required=True, choices=['docx'], help='type of file (example: docx)')
    parser.add_argument('--path', required=True, help='path to file or directory')
    parser.add_argument('--path-to-existing', required=False,
                        help='path to existing questionnaire to add more '
                             'questions')
    parser.add_argument('--path-to-save', help='path to save questionnaire (default value in config.py)',
                        default=QUESTIONNAIRE_FILE)
    parser.add_argument('--save-xlsx', action='store_true', default=False, help='save in xlsx format')
    parser.add_argument('--add-alphabet', action='store_true', default=False,
                        help='add alphabet enumeration to answers')

    args = parser.parse_args()

    if args.filetype == 'docx':  # multiple parsers can be implemented
        Parser = WordParser
    else:
        raise TypeError(f'File format error: {args.filetype} - not implemented')

    files = []
    if args.dir:
        directory = Path(f'./{args.path}')
        pattern = f'*.{args.filetype}'
        files = list(directory.glob(pattern))
    elif args.file:
        print(args.path)
        files = [args.file]

    inst_parser = Parser(
        path_for_parse=files,
        path_to_save=args.path_to_save,
        add_to_quest=args.path_to_existing,
        save_xlsx=args.save_xlsx,
        add_alphabet=args.add_alphabet
    )
    print(f'Parsed questionnaire saved in {inst_parser()}')
