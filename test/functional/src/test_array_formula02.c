/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "xlsxwriter.h"

int main() {

    lxw_workbook  *workbook  = workbook_new("test_array_formula02.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
    lxw_format    *bold      = workbook_add_format(workbook);

    format_set_bold(bold);

    worksheet_write_number(worksheet, 0, 1, 0, NULL);
    worksheet_write_number(worksheet, 1, 1, 0, NULL);
    worksheet_write_number(worksheet, 2, 1, 0, NULL);
    worksheet_write_number(worksheet, 0, 2, 0, NULL);
    worksheet_write_number(worksheet, 1, 2, 0, NULL);
    worksheet_write_number(worksheet, 2, 2, 0, NULL);

    worksheet_write_array_formula(worksheet, RANGE("A1:A3"), "{=SUM(B1:C1*B2:C2)}", bold);

    return workbook_close(workbook);
}
