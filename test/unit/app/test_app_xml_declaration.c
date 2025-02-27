/*
 * Tests for the libxlsxwriter library.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "../ctest.h"
#include "../helper.h"

#include "../../../include/xlsxwriter/app.h"

// Test _xml_declaration().
CTEST(app, xml_declaration) {

    char* got;
    char exp[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    FILE* testfile = lxw_tmpfile(NULL);

    lxw_app *app = lxw_app_new();
    app->file = testfile;

    _app_xml_declaration(app);

    RUN_XLSX_STREQ(exp, got);

    lxw_app_free(app);
}
