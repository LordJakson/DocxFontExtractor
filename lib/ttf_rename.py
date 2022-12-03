import logging

from fontTools import ttLib

logger = logging.getLogger(__name__)

WINDOWS_ENGLISH_IDS = 3, 1, 0x409
MAC_ROMAN_IDS = 1, 0, 0

FAMILY_RELATED_IDS = dict(
        LEGACY_FAMILY=1,
        TRUETYPE_UNIQUE_ID=3,
        FULL_NAME=4,
        POSTSCRIPT_NAME=6,
        PREFERRED_FAMILY=16,
        WWS_FAMILY=21,
        )


def rename_font(font_file, new_name, output_name=None):
    def get_current_family_name(table):
        font_names = []
        for plat_id, enc_id, lang_id in (WINDOWS_ENGLISH_IDS, MAC_ROMAN_IDS):
            for name_id in (FAMILY_RELATED_IDS["PREFERRED_FAMILY"], FAMILY_RELATED_IDS["LEGACY_FAMILY"]):
                family_name_rec = table.getName(nameID=name_id, platformID=plat_id, platEncID=enc_id, langID=lang_id)
                if family_name_rec is not None:
                    font_names.append(family_name_rec.toUnicode())
        if not font_names:
            raise ValueError("family name not found")

        for font_name in font_names:
            logger.info("get_current_family_name: %s", font_name)

        for font_name in font_names:
            if font_name.startswith("___"):
                return font_name
        return font_names[0]

    def rename_record(name_record, family_name, suffix):
        old_family_name = name_record.toUnicode()
        new_family_name = old_family_name.replace(family_name, suffix)

        if old_family_name != new_family_name:
            name_record.string = new_family_name
            return old_family_name, new_family_name

        return old_family_name, "no change"

    def rename_font_family(font, value):
        table = font["name"]

        family_name = get_current_family_name(table)
        logger.info("  Current family name: '%s'", family_name)

        # postcript name can't contain spaces
        ps_family_name = family_name.replace(" ", "")
        ps_value = value.replace(" ", "")
        for rec in table.names:
            name_id = rec.nameID
            if name_id not in FAMILY_RELATED_IDS.values():
                continue
            if name_id == FAMILY_RELATED_IDS["POSTSCRIPT_NAME"]:
                old, new = rename_record(rec, ps_family_name, ps_value)
            elif name_id == FAMILY_RELATED_IDS["TRUETYPE_UNIQUE_ID"]:
                if ps_family_name in rec.toUnicode():
                    old, new = rename_record(rec, ps_family_name, ps_value)
                else:
                    old, new = rename_record(rec, family_name, value)
            else:
                old, new = rename_record(rec, family_name, value)

            logger.info("    %r: '%s' -> '%s'", rec, old, new)

        return family_name

    # main

    if not output_name:
        output_name = font_file

    with ttLib.TTFont(font_file) as f:
        logger.info("Renaming font: '%s'", font_file)
        rename_font_family(f, new_name)
        f.save(output_name)
        logger.info("Saved font: '%s'", output_name)


def display_name_table(font_file):
    with ttLib.TTFont(font_file) as f:
        for rec in f["name"].names:
            print(rec.nameID, rec.toUnicode())
