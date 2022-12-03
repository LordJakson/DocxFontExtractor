import os


def convert_font(odttf_file, key, ttf_file=None, del_odttf=False):
    if not ttf_file:
        ttf_file = os.path.splitext(odttf_file)[0] + '.ttf'

    key = key.replace('-', '').replace('{', '').replace('}', '')

    # Convert to Int reversed
    key_int = [int(key[i - 2:i], 16) for i in range(32, 0, -2)]

    with open(odttf_file, 'rb') as fh_in, open(ttf_file, 'wb') as fh_out:
        cont = fh_in.read()
        fh_out.write(bytes(b ^ key_int[i % len(key_int)] for i, b in enumerate(cont[:32])))
        fh_out.write(cont[32:])

    if del_odttf:
        os.remove(odttf_file)

    return ttf_file
