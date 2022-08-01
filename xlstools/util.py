def col_to_colnum(col):
    """
    Convert an excel column name to a 0-indexed number 'A' = 0, 'B' = 1, ... 'AA' = 26, ... 'AAZ' = 727, ...
    :param col:
    :return:
    """
    if not isinstance(col, str):
        return int(col)
    cols = list(col.upper())
    num = 0
    while len(cols) > 0:
        num *= 26
        c = cols.pop(0)
        num += (ord(c) - ord('A') + 1)
    return num - 1


def colnum_to_col(num):
    """
    Convert a 0-indexed numeric index into an alphabetical column label. 0 = 'A', 1 = 'B'... 26 = 'AA', 27 = 'AB', ...
    :param num:
    :return:
    """
    if isinstance(num, str):
        return num
    col = ''
    num = int(num)
    while 1:
        rad = num % 26
        col = chr(ord('A') + rad) + col
        num //= 26
        num -= 1
        if num < 0:
            return col
