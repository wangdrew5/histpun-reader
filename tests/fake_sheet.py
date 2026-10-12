"""In-memory stand-in for the slice of the xlwings API the scripts use.

Mirrors xlwings' value shapes: a single cell reads as a scalar, a single row
or column as a flat list, a block as a list of lists; .options(ndim=2)
always gives a list of lists. Blank cells read as None.
"""


class _LastCell:
    def __init__(self, row):
        self.row = row


class _UsedRange:
    def __init__(self, sheet):
        self._sheet = sheet

    @property
    def last_cell(self):
        rows = [r for (r, _c), v in self._sheet.cells.items() if v not in (None, "")]
        return _LastCell(max(rows + [1, self._sheet.extra_used_rows]))


class _Range:
    def __init__(self, sheet, a, b, ndim=None):
        self.sheet, self.a, self.b, self.ndim = sheet, a, b, ndim

    def options(self, ndim=None):
        return _Range(self.sheet, self.a, self.b, ndim)

    @property
    def value(self):
        (r1, c1), (r2, c2) = self.a, self.b
        block = [[self.sheet.cells.get((r, c)) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        if self.ndim == 2:
            return block
        if r1 == r2 and c1 == c2:
            return block[0][0]
        if r1 == r2:
            return block[0]
        if c1 == c2:
            return [row[0] for row in block]
        return block

    @value.setter
    def value(self, v):
        if self.a != self.b:
            raise NotImplementedError("fake only supports single-cell writes")
        if self.sheet.writes_before_crash is not None:
            if self.sheet.writes_before_crash == 0:
                raise KeyboardInterrupt("simulated crash mid-write")
            self.sheet.writes_before_crash -= 1
        self.sheet.writes.append((self.a, v))
        self.sheet.cells[self.a] = v


class FakeSheet:
    def __init__(self, headers, rows=(), extra_used_rows=0):
        """headers: list of column names (row 1). rows: list of dicts
        {header: value}, written from row 2 down. extra_used_rows mimics
        formatting that inflates Excel's used range below the data."""
        self.cells = {}
        self.writes = []
        self.extra_used_rows = extra_used_rows
        self.writes_before_crash = None  # set to N to raise on write N+1
        self.headers = list(headers)
        for c, h in enumerate(headers, start=1):
            self.cells[(1, c)] = h
        for i, rec in enumerate(rows, start=2):
            self.set_record(i, rec)

    @property
    def used_range(self):
        return _UsedRange(self)

    @property
    def header_map(self):
        return {h: c for c, h in enumerate(self.headers, start=1)}

    def set_record(self, row, rec):
        for h, v in rec.items():
            self.cells[(row, self.header_map[h])] = v

    def get(self, row, header):
        return self.cells.get((row, self.header_map[header]))

    def insert_rows(self, at, n):
        """Excel 'Insert rows': everything from `at` down moves n rows lower."""
        moved = {}
        for (r, c), v in self.cells.items():
            moved[(r + n if r >= at else r, c)] = v
        self.cells = moved

    def range(self, a, b=None):
        if isinstance(a, str):
            raise NotImplementedError("fake only supports (row, col) addressing")
        return _Range(self, a, b or a)


class FakeBook:
    def __init__(self):
        self.saves = 0

    def save(self):
        self.saves += 1
