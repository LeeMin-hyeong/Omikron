from bisect import bisect_left

from openpyxl.worksheet.worksheet import Worksheet


def named_rows(ws: Worksheet, name_column: int) -> list[tuple[int, str]]:
    """일반 시트에서 빈 셀을 만들지 않고 이름이 저장된 행만 찾는다."""
    return sorted(
        (row, cell.value)
        for (row, col), cell in ws._cells.items()
        if row >= 2 and col == name_column and cell.value is not None
    )


def delete_rows_sparse(ws: Worksheet, rows: list[int]) -> None:
    """delete_rows()와 같은 행 이동을 기존 셀에만 적용한다."""
    if not rows:
        return
    deleted = sorted(set(rows))
    deleted_set = set(deleted)
    cells = {}
    # openpyxl.delete_rows()는 이동 범위의 빈 셀까지 생성하므로 사용하지 않는다.
    # 내부 셀 저장소 접근은 이 모듈에 모은다.
    for (row, col), cell in ws._cells.items():
        if row in deleted_set:
            continue
        new_row = row - bisect_left(deleted, row)
        cell.row = new_row
        cells[new_row, col] = cell
    ws._cells = cells
    ws._current_row = max((row for row, _ in cells), default=0)
