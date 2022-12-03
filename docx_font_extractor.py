import logging
import argparse
from texttable import Texttable

import lib.ttf_rename as ttf_rename
import lib.docx_emb_fonts as docx_emb_fonts
import lib.convert_font as convert_font


def main():
    parser = argparse.ArgumentParser(
            prog='DocxFontExtractor',
            description='Extracts embedded font from a docx file\n'
                        'and converts it to usable TTF files',
            formatter_class=argparse.RawTextHelpFormatter)

    parser.add_argument('filename',
                        help="DocxFile to extract the Fonts from")
    parser.add_argument('-d', '--TargetDir',
                        help="Directory to extract the Fonts to",
                        nargs="?",
                        metavar="fonts",
                        const="fonts",
                        default=None
                        )
    parser.add_argument('-l', '--List',
                        help="List all the embedded Fonts\n"
                             "instead of extracting them",
                        action="store_true"
                        )
    parser.add_argument('-n', '--FontNames',
                        help="List of font names to extract\n"
                             "Regular expression can be used",
                        nargs="+",
                        metavar="'Font Name'",
                        default=None
                        )
    parser.add_argument('-i', '--FontIds',
                        help="List of font ids to extract",
                        nargs="+",
                        metavar="fontX",
                        default=None
                        )
    parser.add_argument('-v', '--Verbose',
                        help="Display logging information's\n"
                             "-v  = INFO\n"
                             "-vv = DEBUG",
                        action='count',
                        default=0)

    args = parser.parse_args()

    log_levels = ["WARNING", "INFO", "DEBUG"]
    log_level = log_levels[min(args.Verbose, len(log_levels) - 1)]
    logging.basicConfig(level=log_level)  # , format="%(message)s"

    with docx_emb_fonts.DocxReader("WordFile.docx") as docx_file:
        target_dir = args.TargetDir

        if args.List:
            font_list = docx_file.get_font_list()
            table = Texttable()
            table.header(["Font name", "Font id"])
            table.set_deco(Texttable.BORDER | Texttable.HEADER | Texttable.VLINES)
            for font in font_list:
                table.add_row([font.font_name + " " + font.font_type, font.font_id])
            print(table.draw())
            exit(0)

        if args.FontNames:
            font_list = docx_file.extract(target_dir, name_type_list=args.FontNames)
        elif args.FontIds:
            font_list = docx_file.extract(target_dir, id_list=args.FontIds)
        else:
            font_list = docx_file.extract_all(target_dir)

        for font in font_list:
            ttf_file = convert_font.convert_font(font.file_name, font.key, del_odttf=True)
            ttf_rename.rename_font(ttf_file, font.font_name)
            # ttf_rename.display_name_table(ttf_file)
            print("Font '{}' extracted to '{}'".format(font.font_name + " " + font.font_type, ttf_file))


if __name__ == '__main__':
    main()
