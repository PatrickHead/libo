/*
 *  Copyright 2019, 2020, 2022, 2025 Patrick Head
 *
 *  This program is free software: you can redistribute it and/or modify it
 *  under the terms of the GNU Lesser General Public License as published by the
 *  Free Software Foundation, either version 3 of the License, or (at your
 *  option) any later version.
 *
 *  This program is distributed in the hope that it will be useful, but WITHOUT
 *  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 *  FITNESS FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public License
 *  for more details.
 *
 *  You should have received a copy of the GNU Lesser General Public License
 *  along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

/**
 *  @file libo.c
 *  @brief  Source code file for libo
 *
 *  A library to aid in manipulating data in Office files.
 *
 *  NOTE:  Currently only XLSX (Excel) files are supported
 */

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <ctype.h>

#include "libo.h"

#if defined(LIBXML_XPATH_ENABLED) && defined(LIBXML_SAX1_ENABLED)
#define XPATH_ENABLED 1  /**<  switches on XPath code when needed   */
#else
#define XPATH_ENABLED 0  /**<  switches off XPath code when needed  */
#endif

static void cell_ref_to_row_col(char *ref, int *row, int *col);
static int is_office(libo *l);
static int is_supported(libo *l);
static libo_type get_type(libo *l);
static void do_indent(FILE *stream, int indent);
static char *find_app_type_name_in_xml(xmlDocPtr doc);
static int count_sheets_in_xml(xmlDocPtr doc);
static int count_sheet_rows_in_xml(xmlDocPtr doc);
static int count_sheet_columns_in_xml(xmlDocPtr doc);
static libo_xl_cell_type string_to_libo_xl_cell_type(char *s);
static int libo_xl_write(libo *l);
static void libo_xl_sheet_count_columns(libo_xl_sheet *xls);
static int libo_xl_themes_write(libo *l);
static int libo_xl_styles_write(libo *l);
static int libo_xl_docprops_write(libo *l);
static int libo_xl_docprops_app_write(libo *l);
static int libo_xl_docprops_core_write(libo *l);
static int libo_xl_xl_rels_write(libo *l);
static int libo_xl__rels_dot_rels_write(libo *l);
static int libo_xl_xl_rels_workbook_rels_write(libo *l);
static int libo_xl_content_types_write(libo *l);
static int libo_xl_workbook_write(libo *l);
static int libo_xl_sheets_write(libo *l);
static int libo_xl_sheet_write(libo *l, int sheet);
static char *strapp(char *s1, char *s2);
static int libo_xl_shared_strings_write(libo *l);
static void libo_xl_renumber_strings(libo *l);
static int libo_xl_shared_strings_write(libo *l);
static void libo_xl_sheet_dimension_add(libo *l, int sheet, char **buf);
static void libo_xl_sheet_sheetviews_add(libo *l, int sheet, char **buf);
static void libo_xl_sheet_formatpr_add(libo *l, int sheet, char **buf);
static void libo_xl_sheet_cols_add(libo *l, int sheet, char **buf);
static void libo_xl_sheet_sheetdata_add(libo *l, int sheet, char **buf);
static void libo_xl_sheet_sheetdata_row_add(libo *l,
                                                int sheet,
                                                int row,
                                                char **buf);
static void libo_xl_sheet_sheetdata_row_col_add(libo *l,
                                                    int sheet,
                                                    int row,
                                                    int col,
                                                    char **buf);
static void libo_xl_sheet_filter_add(libo *l, int sheet, char **buf);
static void libo_xl_strings_count_action(avl_node *n);
static void libo_xl_strings_add_action(avl_node *n);
static char *column_number_to_reference(unsigned int n);
static char *reverse(char *s);
static void libo_xl_row_fill(libo_xl_sheet *sheet, int max_row);
static void libo_xl_col_fill(libo_xl_sheet *sheet, int row, int max_col);
static void string_dumper(avl_node *n);
static void libo_xl_cell_clear(libo_xl_cell *cell);
static libo_xl_column **libo_xl_sheet_columns_create_defaults(libo_xl_sheet *sheet);

static int _strings_count = 0;     /**<  used when counting XL strings  */
static char *_strings_buf = NULL;  /**<  used when accumulating strings
                                         into XML buffer  */


  /**
   *  @fn void libo_init(void)
   *
   *  @brief initialize libo library for later use
   *
   *  @par Parameters
   *  None.
   *
   *  @par Returns
   *  Nothing.
   */

void libo_init(void)
{
  xmlInitParser();
  LIBXML_TEST_VERSION

  return;
}

  /**
   *  @fn void libo_cleanup(void)
   *
   *  @brief called after all use of libo is complete
   *
   *  @par Parameters
   *  None.
   *
   *  @par Returns
   *  Nothing.
   */

void libo_cleanup(void)
{
  xmlCleanupParser();

  return;
}

  /**
   *  @fn libo *libo_new(void)
   *
   *  @brief creates a new @a libo struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo struct
   */

libo *libo_new(void)
{
  libo *l;

  l = (libo *)malloc(sizeof(libo));
  if (!l) return NULL;

  memset(l, 0, sizeof(libo));

  return l;
}

  /**
   *  @fn libo *libo_dup(libo *l)
   *
   *  @brief creates a deep copy of @p l
   *
   *  @param l - pointer to existing libo struct
   *
   *  @return pointer to new @a libo struct
   */

libo *libo_dup(libo *l)
{
  libo *nl = NULL;

  if (!l) goto exit;

  nl = libo_new();
  if (!nl) goto exit;

  if (l->path) nl->path = strdup(l->path);
  nl->type = l->type;
  nl->z = NULL;

  switch (l->type)
  {
    case libo_type_doc:
      nl->doc = libo_doc_dup(l->doc);
      break;

    case libo_type_pp:
      nl->pp = libo_pp_dup(l->pp);
      break;

    case libo_type_xl:
      nl->xl = libo_xl_dup(l->xl);
      break;

    case libo_type_none:
    default:
      break;
  }

exit:
  return nl;
}

  /**
   *  @fn void libo_free(libo *l)
   *
   *  @brief frees all memory allocated to @p l
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_free(libo *l)
{
  if (!l) return;

  if (l->path) free(l->path);
  switch (l->type)
  {
    case libo_type_none: break;
    case libo_type_xl: libo_xl_free(l->xl); break;
    case libo_type_doc: libo_doc_free(l->doc); break;
    case libo_type_pp: libo_pp_free(l->pp); break;
  }

  libo_close(l);

  free(l);

  return;
}

  /**
   *  @fn libo *libo_open(char *path)
   *
   *  @brief creates a new @a libo struct from a file
   *
   *  @param path - name of file to open
   *
   *  @return pointer to new @a libo struct, NULL on error
   */

libo *libo_open(char *path)
{
  libo *l = NULL;
  int err = 0;

  if (!XPATH_ENABLED)
  {
    fprintf(stderr, "XPATH is not enabled in LIBXML2\n");
    return NULL;
  }

  if (!path) return NULL;

  l = libo_new();
  if (!l) return NULL;

  l->path = strdup(path);

  l->z = zip_open(path, ZIP_RDONLY, &err);
  if (!l->z)
  {
    fprintf(stderr, "Can not open '%s', error code is %d\n", path, err);
    return NULL;
  }

  if (!is_office(l))
  {
    libo_free(l);
    fprintf(stderr, "'%s' is does not appear to be an Office document\n", path);
    return NULL;
  }

  l->type = get_type(l);

  if (!is_supported(l))
  {
    libo_free(l);
    fprintf(stderr, "'%s' is not a supported Office document\n", path);
    return NULL;
  }

  switch (l->type)
  {
    case libo_type_none: break;
    case libo_type_xl:
      l->xl = libo_xl_read(l);
      if (!l->xl)
      {
        libo_free(l);
        fprintf(stderr, "Can not establish XL document\n");
        return NULL;
      }
      break;
    case libo_type_doc:
      l->doc = libo_doc_new();
      if (!l->doc)
      {
        libo_free(l);
        fprintf(stderr, "Can not establish DOC document\n");
        return NULL;
      }
      break;
    case libo_type_pp:
      l->pp = libo_pp_new();
      if (!l->pp)
      {
        libo_free(l);
        fprintf(stderr, "Can not establish PP document\n");
        return NULL;
      }
      break;
  }


  return l;
}

  /**
   *  @fn int libo_write(libo *l, char *path)
   *
   *  @brief write libo document to a file
   *
   *  @param l - pointer to existing @a libo struct
   *  @param path - string containing path to file
   *
   *  @return 0 on success, STDIO error on failure
   */

int libo_write(libo *l, char *path)
{
  char *fname = NULL;
  int err = 0;
  int r = 0;

  if (!l) return -1;

  if (path) fname = path;
  else fname = l->path;

  if (!fname) return -1;

  l->z = zip_open(fname, ZIP_CREATE, &err);
  if (!l->z)
  {
    fprintf(stderr, "Can not open '%s', error code is %d\n", path, err);
    return -1;
  }

  switch (l->type)
  {
    case libo_type_xl:
      r = libo_xl_write(l);
      break;
    case libo_type_doc:
      break;
    case libo_type_pp:
      break;
    default: break;
  }

  libo_close(l);

  return r;
}

  /**
   *  @fn void libo_close(libo *l)
   *
   *  @brief closes the open ZIP file associated with @p l
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_close(libo *l)
{
  if (!l) return;

  if (l->z)
  {
    zip_close(l->z);
    l->z = NULL;
  }

  return;
}

  /**
   *  @fn void libo_dump(libo *l, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p l to a file, @b STDOUT by default
   *
   *  @param l - pointer to existing @a libo struct
   *  @param stream - open FILE pointer for writing
   *  @param indent - initial indentation to use for output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_dump(libo *l, FILE *stream, int indent)
{
  if (!l) return;

  if (!stream) stream = stdout;

  do_indent(stream, indent); fprintf(stream, "LIBO:\n");
  indent += 2;
  do_indent(stream, indent); fprintf(stream, "Path: %s\n", l->path);
  do_indent(stream, indent); fprintf(stream, "Type: %s\n", libo_type_to_string(l->type));
  do_indent(stream, indent); fprintf(stream, "z: %p\n", l->z);

  switch (l->type)
  {
    case libo_type_none: break;
    case libo_type_xl:
      libo_xl_dump(l->xl, stream, indent);
      break;
    case libo_type_doc:
      libo_doc_dump(l->doc, stream, indent);
      break;
    case libo_type_pp:
      libo_pp_dump(l->pp, stream, indent);
      break;
  }

  return;
}

  /**
   *  @fn libo_type libo_get_type(libo *l)
   *
   *  @brief returns document type of @p l
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @return @a libo_type
   */

libo_type libo_get_type(libo *l)
{
  if (!l) return libo_type_none;

  return l->type;
}

  /**
   *  @fn void libo_set_type(libo *l, libo_type type)
   *
   *  @brief sets type of @p l to @p type
   *
   *  @param l - pointer to existing @a libo struct
   *  @param type - @a libo_type
   *
   *  @par Returns
   *  Nothing.
   */

void libo_set_type(libo *l, libo_type type)
{
  if (!l) return;

  libo_close(l);

  switch (l->type)
  {
    case libo_type_doc:
      libo_doc_free(l->doc);
      break;

    case libo_type_pp:
      libo_pp_free(l->pp);
      break;

    case libo_type_xl:
      libo_xl_free(l->xl);
      break;

    case libo_type_none:
    default:
      break;
  }

  l->type = type;

  switch (l->type)
  {
    case libo_type_doc:
      l->doc = libo_doc_new();
      break;

    case libo_type_pp:
      l->pp = libo_pp_new();
      break;

    case libo_type_xl:
      l->xl = libo_xl_new();
      break;

    case libo_type_none:
    default:
      l->xl = NULL;
      break;
  }
}

  /**
   *  @fn char *libo_get_path(libo *l)
   *
   *  @brief returns path to @a libo file
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @return string containing path to file
   */

char *libo_get_path(libo *l)
{
  if (!l) return NULL;

  return l->path;
}

  /**
   *  @fn void libo_set_path(libo *l, char *path)
   *
   *  @brief sets path in @p l to @p path
   *
   *  @param l - pointer to existing @a libo struct
   *  @param path - string containing new path
   *
   *  @par Returns
   *  Nothing.
   */

void libo_set_path(libo *l, char *path)
{
  if (!l) return;
  if (l->path) free(l->path);
  l->path = NULL;
  if (path) l->path = strdup(path);
}

  /**
   *  @fn libo_xl *libo_get_xl(libo *l)
   *
   *  @brief returns @a libo_xl from @p l
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @return pointer to @a libo_xl
   */

libo_xl *libo_get_xl(libo *l)
{
  if (!l) return NULL;

  if (l->type == libo_type_xl) return l->xl;

  return NULL;
}

  /**
   *  @fn libo_xl_book *libo_xl_get_book(libo_xl *xl)
   *
   *  @brief returns @a libo_xl_book from @p xl
   *
   *  @param xl - pointer to existing @a libo_xl struct
   *
   *  @return pointer to @a libo_xl_book
   */

libo_xl_book *libo_xl_get_book(libo_xl *xl)
{
  if (!xl) return NULL;

  return xl->book;

  return NULL;
}

  /**
   *  @fn int libo_xl_book_get_sheet_count(libo_xl_book *xlb)
   *
   *  @brief returns number of work sheets in @p xlb
   *
   *  @param xlb - pointer to existing @a libo_xl_book struct
   *
   *  @return count of sheets
   */

int libo_xl_book_get_sheet_count(libo_xl_book *xlb)
{
  if (!xlb) return 0;

  return xlb->n_sheets;
}

  /**
   *  @fn libo_xl_sheet *libo_xl_book_get_sheet(libo_xl_book *xlb, int n)
   *
   *  @brief returns specific work sheet in @p xlb
   *
   *  @param xlb - pointer to existing @a libo_xl_book struct
   *  @param n - index of desired sheet
   *
   *  @return pointer to @a libo_xl_sheet
   */

libo_xl_sheet *libo_xl_book_get_sheet(libo_xl_book *xlb, int n)
{
  if (!xlb) return NULL;

  if (n < 0) return NULL;
  if (n >= xlb->n_sheets) return NULL;

  return xlb->sheet[n];
}

  /**
   *  @fn int libo_xl_sheet_get_row_count(libo_xl_sheet *xls)
   *
   *  @brief returns number of rows in @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @return count of rows
   */

int libo_xl_sheet_get_row_count(libo_xl_sheet *xls)
{
  if (!xls) return 0;

  return xls->n_rows;
}

  /**
   *  @fn int libo_xl_sheet_get_column_count(libo_xl_sheet *xls)
   *
   *  @brief returns number of columns in @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @return count of columns
   */

int libo_xl_sheet_get_column_count(libo_xl_sheet *xls)
{
  if (!xls) return 0;

  return xls->n_cols;
}

  /**
   *  @fn char *libo_xl_sheet_get_name(libo_xl_sheet *xls)
   *
   *  @brief returns title of @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @return string containing title/name of work sheet
   */

char *libo_xl_sheet_get_name(libo_xl_sheet *xls)
{
  if (!xls) return NULL;

  return xls->name;
}

  /**
   *  @fn int libo_xl_sheet_get_id(libo_xl_sheet *xls)
   *
   *  @brief returns identifier of @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @return integer identifier
   */

int libo_xl_sheet_get_id(libo_xl_sheet *xls)
{
  if (!xls) return 0;

  return xls->ID;
}

  /**
   *  @fn char *libo_xl_sheet_get_rid(libo_xl_sheet *xls)
   *
   *  @brief returns relative identifier of @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @return string containing relative identifier
   */

char *libo_xl_sheet_get_rid(libo_xl_sheet *xls)
{
  if (!xls) return NULL;

  return xls->rID;
}

  /**
   *  @fn libo_xl_row *libo_xl_sheet_get_row(libo_xl_sheet *xls, int n)
   *
   *  @brief returns row at index @p n from @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *  @param n - index of row to retrieve
   *
   *  @return pointer to @a libo_xl_row
   */

libo_xl_row *libo_xl_sheet_get_row(libo_xl_sheet *xls, int n)
{
  if (!xls) return NULL;

  if (n < 0) return NULL;
  if (n >= xls->n_rows) return NULL;

  return xls->row[n];

  return NULL;
}

  /**
   *  @fn int libo_xl_row_get_cell_count(libo_xl_row *xlr)
   *
   *  @brief returns number of cells in @p xlr
   *
   *  @param xlr - pointer to existing @a libo_xl_row struct
   *
   *  @return number of cells in @p xlr
   */

int libo_xl_row_get_cell_count(libo_xl_row *xlr)
{
  if (!xlr) return 0;

  return xlr->n_cells;
}

  /**
   *  @fn libo_xl_cell *libo_xl_row_get_cell(libo_xl_row *xlr, int n)
   *
   *  @brief returns cell at index @p n from @p xlr
   *
   *  @param xlr - pointer to existing @a libo_xl_row struct
   *  @param n - index of cell to retrieve
   *
   *  @return pointer to @a libo_xl_cell
   */

libo_xl_cell *libo_xl_row_get_cell(libo_xl_row *xlr, int n)
{
  if (!xlr) return NULL;

  if (n < 0) return NULL;
  if (n >= xlr->n_cells) return NULL;

  return xlr->cell[n];
}

  /**
   *  @fn libo_xl_cell_type libo_xl_cell_get_type(libo_xl_cell *xlc)
   *
   *  @brief returns cell type of @p xlc
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *
   *  @return @a libo_xl_cell_type
   */

libo_xl_cell_type libo_xl_cell_get_type(libo_xl_cell *xlc)
{
  if (!xlc) return libo_xl_cell_type_none;

  return xlc->type;
}

  /**
   *  @fn void libo_xl_cell_set_type(libo_xl_cell *xlc, libo_xl_cell_type type)
   *
   *  @brief sets cell type of @p xlc to @p type
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *  @param type - @a libo_xl_cell_type
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_set_type(libo_xl_cell *xlc, libo_xl_cell_type type)
{
  if (!xlc) return;

  if (xlc->type == libo_xl_cell_type_expression)
  {
    if (xlc->expression.formula) free(xlc->expression.formula);
    if (xlc->expression.value) free(xlc->expression.value);
  }

  memset(&xlc->expression, 0, sizeof(libo_xl_cell_expression));

  xlc->type = type;
}

  /**
   *  @fn char *libo_xl_cell_get_string_value(libo_xl *xl, libo_xl_cell *xlc)
   *
   *  @brief returns string value of @p xlc
   *
   *  NOTE:  @p xl is required, as we need to lookup values in the string dictionary
   *
   *  @param xl - pointer to existing @a libo_xl
   *  @param xlc - pointer to existing @a libo_xl_cell
   *
   *  @return string containing cell's value, NULL if cell type is unknown
   */

char *libo_xl_cell_get_string_value(libo_xl *xl, libo_xl_cell *xlc)
{
  libo_xl_cell_type type;
  libo_xl_cell_expression *expr;
  char *value = NULL;

  if (!xlc) return NULL;

  type = libo_xl_cell_get_type(xlc);

  switch (type)
  {
    case libo_xl_cell_type_none: break;

    case libo_xl_cell_type_reference:
      if (!xl) return NULL;
      value = libo_xl_cell_get_text(xl, xlc);
      if (value) value = strdup(value);
      break;

    case libo_xl_cell_type_expression:
      expr = libo_xl_cell_get_expression(xlc);
      value = strdup(libo_xl_cell_expression_get_formula(expr));
      if (!value)
        value = strdup(libo_xl_cell_expression_get_value(expr));
      if (!value)
        value = strdup("");
      break;

    case libo_xl_cell_type_number:
      value = (char *)malloc(50);
      memset(value, 0, 50);
      sprintf(value, "%g", libo_xl_cell_get_number(xlc));
      break;
  }

  return value;
}

  /**
   *  @fn int libo_xl_cell_get_reference(libo_xl_cell *xlc)
   *
   *  @brief returns reference id in @p xlc
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *
   *  @return reference id of @p xld, 0 if cell type is not a reference
   */

int libo_xl_cell_get_reference(libo_xl_cell *xlc)
{
  if (!xlc) return 0;
  if (xlc->type != libo_xl_cell_type_reference) return 0;

  return xlc->reference;
}

  /**
   *  @fn void libo_xl_cell_set_reference(libo_xl_cell *xlc, int reference)
   *
   *  @brief sets reference value in @p xlc to @p reference
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *  @param reference - new reference value
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_set_reference(libo_xl_cell *xlc, int reference)
{
  if (!xlc) return;
  if (xlc->type != libo_xl_cell_type_reference) return;

  xlc->reference = reference;
}

  /**
   *  @fn char *libo_xl_cell_get_text(libo_xl *xl, libo_xl_cell *xlc)
   *
   *  @brief returns string value of @p xlc
   *
   *  NOTE:  @p xl is required, as we need to lookup values in the string dictionary
   *
   *  @param xl - pointer to existing @a libo_xl
   *  @param xlc - pointer to existing @a libo_xl_cell
   *
   *  @return string value of @p xlc
   */

char *libo_xl_cell_get_text(libo_xl *xl, libo_xl_cell *xlc)
{
  string *str;
  char *value = NULL;

  if (!xl) goto exit;
  if (!xlc) goto exit;
  if (xlc->type != libo_xl_cell_type_reference) goto exit;

  str = strings_find_by_id(xl->strings, xlc->reference);
  if (str) value = str->text;

exit:
  return value;
}

  /**
   *  @fn void libo_xl_cell_set_text(libo_xl *xl, libo_xl_cell *xlc, char *text)
   *
   *  @brief sets text value in @p xlc to @p text
   *
   *  @param xl - pointer to existing @a libo_xl
   *  @param xlc - pointer to existing @a libo_xl_cell
   *  @param text - new string value for text
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_set_text(libo_xl *xl, libo_xl_cell *xlc, char *text)
{
  string *str;

  if (!xl) return;
  if (!xlc) return;

  libo_xl_cell_clear(xlc);

  libo_xl_cell_set_type(xlc, libo_xl_cell_type_reference);

  str = strings_find_by_text(xl->strings, text);
  if (!str)
  {
    str = string_new_with_values(text, 0);
    if (str) strings_add(xl->strings, str);
    str = strings_find_by_text(xl->strings, text);
  }

  if (str) libo_xl_cell_set_reference(xlc, str->id);
}

  /**
   *  @fn libo_xl_cell_expression *libo_xl_cell_get_expression(libo_xl_cell *xlc)
   *
   *  @brief returns cell expression from @p xlc
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *
   *  @return pointer to @a libo_xl_cell_expression, NULL if cell type is not an expression
   */

libo_xl_cell_expression *libo_xl_cell_get_expression(libo_xl_cell *xlc)
{
  if (!xlc) return NULL;
  if (xlc->type != libo_xl_cell_type_expression) return NULL;

  return &xlc->expression;
}

  /**
   *  @fn void libo_xl_cell_set_expression(libo_xl_cell *xlc, libo_xl_cell_expression *xlce)
   *
   *  @brief sets @p xlc expression values to those in @p xlce
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *  @param xlce - pointer to existing @a libo_xl_cell_expression
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_set_expression(libo_xl_cell *xlc, libo_xl_cell_expression *xlce)
{
  if (!xlc || !xlce) return;

  libo_xl_cell_clear(xlc);

  libo_xl_cell_set_type(xlc, libo_xl_cell_type_expression);

  if (xlce->formula) libo_xl_cell_expression_set_formula(&xlc->expression, xlce->formula);
  if (xlce->value) libo_xl_cell_expression_set_value(&xlc->expression, xlce->value);
}

  /**
   *  @fn char *libo_xl_cell_expression_get_formula(libo_xl_cell_expression *xlce)
   *
   *  @brief returns formula from @p xlce
   *
   *  @param xlce - pointer to existing @a libo_xl_cell_expression
   *
   *  @return string containing cell's formula
   */

char *libo_xl_cell_expression_get_formula(libo_xl_cell_expression *xlce)
{
  if (!xlce) return NULL;

  return xlce->formula;
}

  /**
   *  @fn void libo_xl_cell_expression_set_formula(libo_xl_cell_expression *xlce, char *formula)
   *
   *  @brief sets formula in @p xlce to @p formula
   *
   *  @param xlce - pointer to existing @a libo_xl_cell_expression
   *  @param formula - string containing new formula
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_expression_set_formula(libo_xl_cell_expression *xlce,
                                         char *formula)
{
  char *old_formula = NULL;

  if (!xlce || !formula) return;

  if (xlce->formula) old_formula = xlce->formula;

  xlce->formula = NULL;

  xlce->formula = strdup(formula);

  if (old_formula) free(old_formula);
}

  /**
   *  @fn char *libo_xl_cell_expression_get_value(libo_xl_cell_expression *xlce)
   *
   *  @brief returns value from @p xlce
   *
   *  @param xlce - pointer to existing @a libo_xl_cell_expression
   *
   *  @return string containing cell's calculated value
   */

char *libo_xl_cell_expression_get_value(libo_xl_cell_expression *xlce)
{
  if (!xlce) return NULL;

  return xlce->value;
}

  /**
   *  @fn void libo_xl_cell_expression_set_value(libo_xl_cell_expression *xlce,
   *                                             char *value)
   *
   *  @brief sets value in @p xlce to @p value
   *
   *  @param xlce - pointer to existing @a libo_xl_cell_expression
   *  @param value - string containing new value
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_expression_set_value(libo_xl_cell_expression *xlce,
                                       char *value)
{
  char *old_value = NULL;

  if (!xlce || !value) return;

  if (xlce->value) old_value = xlce->value;
  xlce->value = NULL;

  if (value) xlce->value = strdup(value);

  if (old_value) free(old_value);
}

  /**
   *  @fn double libo_xl_cell_get_number(libo_xl_cell *xlc)
   *
   *  @brief returns direct value from @p xlc
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *
   *  @return direct numeric value of @p xlc
   */

double libo_xl_cell_get_number(libo_xl_cell *xlc)
{
  if (!xlc) return 0;
  if (xlc->type != libo_xl_cell_type_number) return 0;

  return xlc->number;
}

  /**
   *  @fn void libo_xl_cell_set_number(libo_xl_cell *xlc, double number)
   *
   *  @brief sets number value in @p xlc to @p number
   *
   *  @param xlc - pointer to existing @a libo_xl_cell
   *  @param number - new number value
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_set_number(libo_xl_cell *xlc, double number)
{
  if (!xlc) return;

  libo_xl_cell_clear(xlc);

  libo_xl_cell_set_type(xlc, libo_xl_cell_type_number);

  xlc->number = number;
}

  /**
   *  @fn libo_doc *libo_get_doc(libo *l)
   *
   *  @brief returns Word document from @p l
   *
   *  @param l - pointer to existing @a libo
   *
   *  @return pointer to @a libo_doc, NULL if document type is not a Word document
   */

libo_doc *libo_get_doc(libo *l)
{
  if (!l) return NULL;

  if (l->type == libo_type_doc) return l->doc;

  return NULL;
}

  /**
   *  @fn libo_pp *libo_get_pp(libo *l)
   *
   *  @brief returns PowerPoint document from @p l
   *
   *  @param l - pointer to existing @a libo
   *
   *  @return pointer to @a libo_pp, NULL if document type is not a PowerPoint document
   */

libo_pp *libo_get_pp(libo *l)
{
  if (!l) return NULL;

  if (l->type == libo_type_pp) return l->pp;

  return NULL;
}

  /**
   *  @fn libo_xl *libo_xl_new(void)
   *
   *  @brief creates a new @a libo_xl struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to @a libo_xl
   */

libo_xl *libo_xl_new(void)
{
  libo_xl *xl;

  xl = (libo_xl *)malloc(sizeof(libo_xl));
  if (!xl) return NULL;
  memset(xl, 0, sizeof(libo_xl));

  xl->book = libo_xl_book_new();
  xl->strings = strings_new();

  return xl;
}

  /**
   *  @fn libo_xl *libo_xl_dup(libo_xl *xl)
   *
   *  @brief creates a deep copy of @p xl
   *
   *  @param xl - pointer to existing @a libo_xlstruct
   *
   *  @return pointer to new @a libo_xl struct
   */

libo_xl *libo_xl_dup(libo_xl *xl)
{
  libo_xl *nxl = NULL;

  if (!xl) goto exit;

  nxl = libo_xl_new();
  if (!nxl) goto exit;

  if (xl->book) nxl->book = libo_xl_book_dup(xl->book);
  if (xl->strings) nxl->strings = strings_dup(xl->strings);

exit:
  return nxl;
}

  /**
   *  @fn void libo_xl_free(libo_xl *xl)
   *
   *  @brief frees all memory allocated to @p xl
   *
   *  @param xl - pointer to existing @a libo_xl struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_free(libo_xl *xl)
{
  if (!xl) return;

  if (xl->book) libo_xl_book_free(xl->book);
  if (xl->strings) strings_free(xl->strings);

  free(xl);

  return;
}

  /**
   *  @fn void libo_xl_sheet_set_default_row_height(libo_xl_sheet *sheet,
   *                                                double default_row_height)
   *
   *  @brief sets default row height of @p sheet
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *  @param default_row_height
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_set_default_row_height(libo_xl_sheet *sheet,
                                          double default_row_height)
{
  if (!sheet) return;

  sheet->default_row_height = default_row_height;
}

  /**
   *  @fn libo_xl_freeze *libo_xl_sheet_get_freeze(libo_xl_sheet *sheet)
   *
   *  @brief gets row/column freeze from @p sheet, if any
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *
   *  @return pointer to @a libo_xl_freeze struct, or NULL if none.
   */

libo_xl_freeze *libo_xl_sheet_get_freeze(libo_xl_sheet *sheet)
{
  if (!sheet) return NULL;

  return &(sheet->freeze);
}

  /**
   *  @fn void libo_xl_sheet_set_freeze(libo_xl_sheet *sheet,
   *                                    libo_xl_freeze_type type,
   *                                    int n)
   *
   *  @brief sets row/column freeze for @p sheet
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *  @param type - @a libo_xl_freeze_type (top or left)
   *  @param n - number of rows or columns to freeze
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_set_freeze(libo_xl_sheet *sheet,
                              libo_xl_freeze_type type,
                              int n)
{
  if (!sheet) return;

  sheet->freeze.type = type;
  sheet->freeze.n = n;
}

  /**
   *  @fn void libo_xl_sheet_add_filter(libo_xl_sheet *sheet,
   *                                    unsigned int first_column,
   *                                    unsigned int last_column)
   *
   *  @brief adds columns filter to XL worksheet
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *  @param first_column
   *  @param last_column
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_add_filter(libo_xl_sheet *sheet,
                              unsigned int first_column,
                              unsigned int last_column)
{
  if (!sheet) return;

  if (sheet->filter) libo_xl_filter_free(sheet->filter);
  sheet->filter = libo_xl_filter_new_with_values(first_column, last_column);
}

  /**
   *  @fn void libo_xl_sheet_remove_filter(libo_xl_sheet *sheet)
   *
   *  @brief removes filter from XL worksheet
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_remove_filter(libo_xl_sheet *sheet)
{
  if (!sheet) return;

  libo_xl_filter_free(sheet->filter);
  sheet->filter = NULL;
}

  /**
   *  @fn libo_doc *libo_doc_new(void)
   *
   *  @brief creates a new @a libo_doc struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to @a libo_doc
   */

libo_doc *libo_doc_new(void)
{
  libo_doc *doc;

  doc = (libo_doc *)malloc(sizeof(libo_doc));
  if (!doc) return NULL;
  memset(doc, 0, sizeof(libo_doc));

  return doc;
}

  /**
   *  @fn libo_doc *libo_doc_dup(libo_doc *doc)
   *
   *  @brief creates a deep copy of @p doc
   *
   *  @param doc - pointer to existing @a libo_doc struct
   *
   *  @return pointer to @a libo_doc
   */

libo_doc *libo_doc_dup(libo_doc *doc)
{
  libo_doc *ndoc = NULL;

  if (!doc) goto exit;

  ndoc = libo_doc_new();
  if (!ndoc) goto exit;

exit:
  return ndoc;
}

  /**
   *  @fn void libo_doc_free(libo_doc *doc)
   *
   *  @brief frees all memory allocated to @p doc
   *
   *  @param doc - pointer to existing @a libo_doc struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_doc_free(libo_doc *doc)
{
  if (!doc) return;

  free(doc);

  return;
}

  /**
   *  @fn libo_pp *libo_pp_new(void)
   *
   *  @brief creates a new @a libo_pp struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to @a libo_pp
   */

libo_pp *libo_pp_new(void)
{
  libo_pp *pp;

  pp = (libo_pp *)malloc(sizeof(libo_pp));
  if (!pp) return NULL;
  memset(pp, 0, sizeof(libo_pp));

  return pp;
}

  /**
   *  @fn libo_pp *libo_pp_dup(libo_pp *pp)
   *
   *  @brief creates a deep copy of @p pp
   *
   *  @param pp - pointer to existing @a libo_pp struct
   *
   *  @return pointer to @a libo_pp
   */

libo_pp *libo_pp_dup(libo_pp *pp)
{
  libo_pp *npp = NULL;

  if (!pp) goto exit;

  npp = libo_pp_new();
  if (!npp) goto exit;

exit:
  return npp;
}

  /**
   *  @fn void libo_pp_free(libo_pp *pp)
   *
   *  @brief frees all memory allocated to @p pp
   *
   *  @param pp - pointer to existing @a libo_pp struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_pp_free(libo_pp *pp)
{
  if (!pp) return;

  free(pp);

  return;
}

  /**
   *  @fn char *libo_type_to_string(libo_type lt)
   *
   *  @brief returns string representation of @p lt
   *
   *  @param lt - @a libo_type
   *
   *  @return string representation of @p lt
   */

char *libo_type_to_string(libo_type lt)
{
  switch (lt)
  {
    case libo_type_none: return "[UNKNOWN]";
    case libo_type_xl: return "Microsoft Excel";
    case libo_type_doc: return "Microsoft Word";
    case libo_type_pp: return "Microsoft PowerPoint";
  }

  return libo_type_none;
}

  /**
   *  @fn libo_xl_strings *libo_xl_strings_read(libo *l)
   *
   *  @brief creates new @a libo_xl_strings struct from file
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @return pointer to new and filled @a libo_xl_strings struct
   */

strings *libo_xl_strings_read(libo *l)
{
  char *strings_file_name = "xl/sharedStrings.xml";
  zip_stat_t stat;
  zip_file_t *zf = NULL;
  xmlDocPtr doc = NULL;
  xmlNodePtr node = NULL;
  char *buf = NULL;
  int len;
  strings *strings;
  string *str;

  if (!l) return NULL;
  if (!l->z) return NULL;

    // open xl/shareStrings.xml

  if (zip_stat(l->z, strings_file_name, 0, &stat)) return NULL;
  if (!((stat.valid & ZIP_STAT_NAME) && (stat.valid & ZIP_STAT_SIZE))) return NULL;

  if (strcmp(strings_file_name, stat.name)) return NULL;

  len = stat.size;

  zf = zip_fopen(l->z, strings_file_name, 0);
  if (!zf)
  {
    fprintf(stderr, "Can not open '%s'\n", strings_file_name); fflush(stderr);
    return NULL;
  }

  buf = (char *)malloc(len+1);
  if (!buf)
  {
    zip_fclose(zf);
    return NULL;
  }
  memset(buf, 0, len+1);

  zip_fread(zf, buf, len);

  zip_fclose(zf);

  doc = xmlParseMemory(buf, len);
  if (!doc)
  {
    fprintf(stderr, "Failed to parse app.xml\n"); fflush(stderr);
    return NULL;
  }
  free(buf);

    // Fill libo_xl_strings structure with strings from XML

  strings = strings_new();
  if (!strings)
  {
    xmlFreeDoc(doc);
    return NULL;
  }

  node = xmlDocGetRootElement(doc);
  if (!strcmp((char *)node->name, "sst"))
  {
    node = node->children;
    while (node)
    {
      if (!strcmp((char *)node->name, "si"))
      {
        buf = (char *)xmlNodeGetContent(node);
        str = string_new_with_values(buf, 0);
        if (str) strings_add(strings, str);
      }
      node = node->next;
    }
  }

  xmlFreeDoc(doc);

  return strings;
}

static int _strings_indent = 0;    /**<  Used by string_dumper()  */
static FILE *_dumper_file = NULL;  /**<  Used by dumpers  */

  /**
   *  @fn void libo_xl_strings_dump(strings *strs, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p lxs to @p stream, default is STDOUT
   *
   *  @param strs - pointer to @a strings struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_strings_dump(strings *strs, FILE *stream, int indent)
{
  if (!strs) return;

  if (!stream) stream = stdout;

  do_indent(stream, indent); fprintf(stream,
                                     "Shared strings (%d):\n",
                                     strs->last_id);
  indent += 2;

  _dumper_file = stream;
  _strings_indent = indent;

  avl_walk(strs->id_root, string_id, string_dumper);

  return;
}

libo_xl *__xl__ = NULL ;  /**<  CHEAT -- short cut to parent @a libo_xl for various dumps  */

  /**
   *  @fn void libo_xl_dump(libo_xl *xl, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p xl to @p stream, default is STDOUT
   *
   *  @param xl - pointer to existing @a libo_xl struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_dump(libo_xl *xl, FILE *stream, int indent)
{
  if (!xl) return;
  if (!stream) return;

  __xl__ = xl;

  do_indent(stream, indent); fprintf(stream, "LIBO_XL:\n");
  indent += 2;
  if (xl->book) libo_xl_book_dump(xl->book, stream, indent);
  if (xl->strings) libo_xl_strings_dump(xl->strings, stream, indent);

  return;
}

  /**
   *  @fn void libo_doc_dump(libo_doc *doc, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p doc to @p stream, default is STDOUT
   *
   *  @param doc - pointer to existing @a libo_doc struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_doc_dump(libo_doc *doc, FILE *stream, int indent)
{
  if (!doc) return;
  if (!stream) return;

  do_indent(stream, indent); fprintf(stream, "LIBO_DOC:\n");
  return;
}

  /**
   *  @fn void libo_pp_dump(libo_pp *pp, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p pp to @p stream, default is STDOUT
   *
   *  @param pp - pointer to existing @a libo_pp struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_pp_dump(libo_pp *pp, FILE *stream, int indent)
{
  if (!pp) return;
  if (!stream) return;

  do_indent(stream, indent); fprintf(stream, "LIBO_PP:\n");
  return;
}

  /**
   *  @fn libo_xl_book *libo_xl_book_new(void)
   *
   *  @brief returns new @a libo_xl_book struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo_xl_book struct
   */

libo_xl_book *libo_xl_book_new(void)
{
  libo_xl_book *book;

  book = (libo_xl_book *)malloc(sizeof(libo_xl_book));
  if (!book) return NULL;
  memset(book, 0, sizeof(libo_xl_book));

  return book;
}

  /**
   *  @fn libo_xl_book *libo_xl_book_dup(libo_xl_book *book)
   *
   *  @brief creates a deep copy of @p book
   *
   *  @param book - pointer to existing @a libo_xl_book struct
   *
   *  @return pointer to new @a libo_xl_book struct
   */

libo_xl_book *libo_xl_book_dup(libo_xl_book *book)
{
  libo_xl_book *nbook = NULL;
  int i;

  if (!book) goto exit;

  nbook = libo_xl_book_new();
  if (!nbook) goto exit;

  nbook->n_sheets = book->n_sheets;

  for (i = 0; i < book->n_sheets; i++)
    libo_xl_book_add(nbook, book->sheet[i]);

exit:
  return nbook;
}

  /**
   *  @fn void libo_xl_book_free(libo_xl_book *book)
   *
   *  @brief frees all memory allocated to @p book
   *
   *  @param book - pointer to existing @a libo_xl_book struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_book_free(libo_xl_book *book)
{
  int i;

  if (!book) return;

  for (i = 0; i < book->n_sheets; i++)
    libo_xl_sheet_free(book->sheet[i]);

  free(book);

  return;
}

  /**
   *  @fn void libo_xl_book_add(libo_xl_book *xlb, libo_xl_sheet *xls)
   *
   *  @brief add @p xls to @p xlb
   *
   *  @param xlb - pointer to existing @a libo_xl_book struct
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_book_add(libo_xl_book *xlb, libo_xl_sheet *xls)
{
  libo_xl_sheet **tmp = NULL;
  char rid_str[128];

  if (!xlb || !xls) return;

  memset(rid_str, 0, 128);

  tmp = realloc(xlb->sheet, sizeof(libo_xl_sheet *) * (xlb->n_sheets + 1));
  if (!tmp) return;

  xlb->sheet = tmp;

  xlb->sheet[xlb->n_sheets] = libo_xl_sheet_dup(xls);

  xlb->sheet[xlb->n_sheets]->ID = xlb->n_sheets + 1;

  sprintf(rid_str, "rId%d", xlb->n_sheets + 4);
  if (xlb->sheet[xlb->n_sheets]->rID)
    free(xlb->sheet[xlb->n_sheets]->rID);
  xlb->sheet[xlb->n_sheets]->rID = strdup(rid_str);

  ++xlb->n_sheets;
}

  /**
   *  @fn libo_xl_sheet *libo_xl_sheet_new(void)
   *
   *  @brief returns new @a libo_xl_sheet struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo_xl_sheet struct
   */

libo_xl_sheet *libo_xl_sheet_new(void)
{
  libo_xl_sheet *sheet;

  sheet = (libo_xl_sheet *)malloc(sizeof(libo_xl_sheet));
  if (!sheet) return NULL;
  memset(sheet, 0, sizeof(libo_xl_sheet));

  sheet->default_row_height = 14.4;

  return sheet;
}

  /**
   *  @fn libo_xl_sheet *libo_xl_sheet_dup(libo_xl_sheet *sheet)
   *
   *  @brief makes deep copy of @p sheet
   *
   *  @param sheet - pointer to existing @a libo_xml_sheet struct
   *
   *  @return pointer to new @a libo_xl_sheet struct
   */

libo_xl_sheet *libo_xl_sheet_dup(libo_xl_sheet *sheet)
{
  libo_xl_sheet *nsheet = NULL;
  int i;

  if (!sheet) goto exit;

  nsheet = libo_xl_sheet_new();
  if (!nsheet) goto exit;

  nsheet->n_cols = sheet->n_cols;
  nsheet->default_row_height = sheet->default_row_height;

  if (sheet->name) nsheet->name = strdup(sheet->name);

  for (i = 0; i < sheet->n_rows; i++)
    libo_xl_sheet_add(nsheet, sheet->row[i]);

exit:
  return nsheet;
}

  /**
   *  @fn void libo_xl_sheet_free(libo_xl_sheet *sheet)
   *
   *  @brief frees all memory allocated to @p sheet
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_free(libo_xl_sheet *sheet)
{
  int i;

  if (!sheet) return;

  for (i = 0; i < sheet->n_rows; i++)
    libo_xl_row_free(sheet->row[i]);

  free(sheet);

  return;
}

  /**
   *  @fn void libo_xl_sheet_set_name(libo_xl_sheet *xls, char *name);
   *
   *  @brief sets name value for @p xls to @p name
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *  @param name - string containing new name value
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_set_name(libo_xl_sheet *xls, char *name)
{
  if (!xls) return;

  if (xls->name) free(xls->name);
  xls->name = NULL;
  if (name) xls->name = strdup(name);
}

  /**
   *  @fn void libo_xl_sheet_set_id(libo_xl_sheet *xls, int id)
   *
   *  @brief sets ID value for @p xls to @p id
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *  @param id - string containing new ID value
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_set_id(libo_xl_sheet *xls, int id)
{
  if (!xls) return;

  xls->ID = id;
}

  /**
   *  @fn void libo_xl_sheet_set_rid(libo_xl_sheet *xls, char *rid)
   *
   *  @brief sets rID value for @p xls to @p rid
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *  @param rid - string containing new rID value
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_set_rid(libo_xl_sheet *xls, char *rid)
{
  if (!xls) return;

  if (xls->rID) free(xls->rID);
  xls->rID = NULL;
  if (rid) xls->rID = strdup(rid);
}

  /**
   *  @fn void libo_xl_sheet_add(libo_xl_sheet *xls, libo_xl_row *xlr)
   *
   *  @brief add @p xlr to @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *  @param xlr - pointer to existing @a libo_xl_row struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_add(libo_xl_sheet *xls, libo_xl_row *xlr)
{
  libo_xl_row **tmp = NULL;

  if (!xls || !xlr) return;

  tmp = realloc(xls->row, sizeof(libo_xl_row *) * (xls->n_rows + 1));
  if (!tmp) return;

  xls->row = tmp;

  xls->row[xls->n_rows] = libo_xl_row_dup(xlr);

  ++xls->n_rows;
}

  /**
   *  @fn libo_xl_row *libo_xl_row_new(void)
   *
   *  @brief returns new @a libo_xl_row struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo_xl_row struct
   */

libo_xl_row *libo_xl_row_new(void)
{
  libo_xl_row *row;

  row = (libo_xl_row *)malloc(sizeof(libo_xl_row));
  if (!row) return NULL;
  memset(row, 0, sizeof(libo_xl_row));

  return row;
}

  /**
   *  @fn libo_xl_row *libo_xl_row_dup(libo_xl_row *row)
   *
   *  @brief returns deep copy of @p row
   *
   *  @param row - pointer to existing @a libo_xl_row struct
   *
   *  @return pointer to new @a libo_xl_row struct
   */

libo_xl_row *libo_xl_row_dup(libo_xl_row *row)
{
  libo_xl_row *nrow = NULL;
  int i;

  if (!row) goto exit;

  nrow = libo_xl_row_new();
  if (!nrow) goto exit;

  for (i = 0; i < row->n_cells; i++)
    libo_xl_row_add(nrow, row->cell[i]);

exit:
  return nrow;
}

  /**
   *  @fn void libo_xl_row_free(libo_xl_row *row)
   *
   *  @brief frees all memory allocated to @p row
   *
   *  @param row - pointer to existing @a libo_xl_row struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_row_free(libo_xl_row *row)
{
  int i;

  if (!row) return;

  for (i = 0; i < row->n_cells; i++)
    libo_xl_cell_free(row->cell[i]);

  free(row);

  return;
}

  /**
   *  @fn void libo_xl_row_add(libo_xl_row *xlr, libo_xl_cell *xlc)
   *
   *  @brief add @p xlc to @p xlr
   *
   *  @param xlr - pointer to existing @a libo_xl_row struct
   *  @param xlc - pointer to existing @a libo_xl_cell struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_row_add(libo_xl_row *xlr, libo_xl_cell *xlc)
{
  libo_xl_cell **tmp = NULL;

  if (!xlr || !xlc) return;

  tmp = realloc(xlr->cell, sizeof(libo_xl_cell *) * (xlr->n_cells + 1));
  if (!tmp) return;

  xlr->cell = tmp;

  xlr->cell[xlr->n_cells] = libo_xl_cell_dup(xlc);

  ++xlr->n_cells;
}

  /**
   *  @fn libo_xl_cell *libo_xl_cell_new(void)
   *
   *  @brief returns new @a libo_xl_cell struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo_xl_cell struct
   */

libo_xl_cell *libo_xl_cell_new(void)
{
  libo_xl_cell *cell;

  cell = (libo_xl_cell *)malloc(sizeof(libo_xl_cell));
  if (cell) memset(cell, 0, sizeof(libo_xl_cell));

  return cell;
}

  /**
   *  @fn libo_xl_cell *libo_xl_cell_dup(libo_xl_cell *cell)
   *
   *  @brief returns deep copy of @p cell
   *
   *  @param cell - pointer to existing @a libo_xl_cell struct
   *
   *  @return pointer to new @a libo_xl_cell struct
   */

libo_xl_cell *libo_xl_cell_dup(libo_xl_cell *cell)
{
  libo_xl_cell *ncell = NULL;

  if (!cell) goto exit;

  ncell = libo_xl_cell_new();
  if (!ncell) goto exit;

  memcpy(ncell, cell, sizeof(libo_xl_cell));

  if (cell->type == libo_xl_cell_type_expression)
  {
    if (cell->expression.formula)
      libo_xl_cell_expression_set_formula(&cell->expression,
                                          cell->expression.formula);
    if (cell->expression.value)
      libo_xl_cell_expression_set_value(&cell->expression,
                                        cell->expression.value);
  }

exit:
  return ncell;
}

  /**
   *  @fn libo_xl_cell *libo_xl_cell_create(libo_xl_sheet *sheet,
   *                                        int row,
   *                                        int col)
   *
   *  @brief returns deep copy of @p cell
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *  @param row - row index to into which to add the cell
   *  @param col - col index of cell to create
   *
   *  @return pointer to new @a libo_xl_cell struct
   */

libo_xl_cell *libo_xl_cell_create(libo_xl_sheet *sheet, int row, int col)
{
  if (!sheet) return NULL;
  if (row < 0) return NULL;
  if (col < 0) return NULL;

  if ((row < sheet->n_rows) && (col < sheet->row[row]->n_cells))
    return sheet->row[row]->cell[col];

  libo_xl_col_fill(sheet, row, col);

  return NULL;
}

  /**
   *  @fn void libo_xl_cell_free(libo_xl_cell *cell)
   *
   *  @brief frees all memory allocated to @p cell
   *
   *  @param cell - pointer to existing @a libo_xl_cell struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_free(libo_xl_cell *cell)
{
  if (!cell) return;

  free(cell);

  return;
}

  /**
   *  @fn libo_xl *libo_xl_read(libo *l)
   *
   *  @brief reads entire contents of Excel parts of document file
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  NOTE:  @p l must already have read ZIP file
   *
   *  @return pointer to new @a lib_xl struct
   */

libo_xl *libo_xl_read(libo *l)
{
  libo_xl *xl;

  if (!l) return NULL;

  xl = libo_xl_new();
  if (!xl) return NULL;

  xl->book = libo_xl_book_read(l);
  xl->strings = libo_xl_strings_read(l);

  return xl;
}

  /**
   *  @fn libo_xl_book *libo_xl_book_read(libo *l)
   *
   *  @brief reads work book contents of Excel parts of document file
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  NOTE:  @p l must already have read ZIP file
   *
   *  @return pointer to new @a lib_xl struct
   */

libo_xl_book *libo_xl_book_read(libo *l)
{
  libo_xl_book *book;
  zip_file_t *zf;
  int len;
  char *workbook_file_name = "xl/workbook.xml";
  char *buf;
  zip_stat_t stat;
  xmlDocPtr doc = NULL;
  int i;

  book = libo_xl_book_new();
  if (!book) return NULL;

    // open xl/workbook.xml

  if (zip_stat(l->z, workbook_file_name, 0, &stat)) return NULL;
  if (!((stat.valid & ZIP_STAT_NAME) && (stat.valid & ZIP_STAT_SIZE))) return NULL;

  if (strcmp(workbook_file_name, stat.name)) return NULL;

  len = stat.size;

  zf = zip_fopen(l->z, workbook_file_name, 0);
  if (!zf)
  {
    fprintf(stderr, "Can not open '%s'\n", workbook_file_name); fflush(stderr);
    return NULL;
  }

  buf = (char *)malloc(len+1);
  if (!buf)
  {
    zip_fclose(zf);
    return NULL;
  }
  memset(buf, 0, len+1);

  zip_fread(zf, buf, len);

  zip_fclose(zf);

  doc = xmlParseMemory(buf, len);
  if (!doc)
  {
    fprintf(stderr, "Failed to parse '%s'\n", workbook_file_name); fflush(stderr);
    return NULL;
  }
  free(buf);

  book->n_sheets = count_sheets_in_xml(doc);

  book->sheet = (libo_xl_sheet **)malloc(sizeof(libo_xl_sheet *) * book->n_sheets);
  if (!book->sheet)
  {
    xmlFreeDoc(doc);
    return NULL;
  }

  for (i = 0; i < book->n_sheets; i++)
  {
    book->sheet[i] = libo_xl_sheet_meta_read(doc, i);
    libo_xl_sheet_read(l, book->sheet[i], i);
  }

  xmlFreeDoc(doc);

  return book;
}

  /**
   *  @fn void libo_xl_book_dump(libo_xl_book *lxb, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p lxb to @p stream, default is STDOUT
   *
   *  @param lxb - pointer to existing @a libo_xl_box struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_book_dump(libo_xl_book *lxb, FILE *stream, int indent)
{
  int i;

  if (!lxb) return;

  if (!stream) stream = stdout;

  do_indent(stream, indent); fprintf(stream, "Book:\n");
  indent += 2;

  do_indent(stream, indent); fprintf(stream, "Sheets (%d):\n", lxb->n_sheets);

  indent += 2;

  for (i = 0; i < lxb->n_sheets; i++)
    libo_xl_sheet_dump(lxb->sheet[i], stream, indent);

  return;
}

  /**
   *  @fn void libo_xl_sheet_dump(libo_xl_sheet *lxs, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p lxs to @p stream, default is STDOUT
   *
   *  @param lxs - pointer to existing @a libo_xl_sheet struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_dump(libo_xl_sheet *lxs, FILE *stream, int indent)
{
  int i;

  if (!lxs) return;

  do_indent(stream, indent); fprintf(stream, "Sheet:\n");

  indent += 2;

  do_indent(stream, indent); fprintf(stream, "Name: %s\n", lxs->name);
  do_indent(stream, indent); fprintf(stream, "ID: %d\n", lxs->ID);
  do_indent(stream, indent); fprintf(stream, "rID: %s\n", lxs->rID);

  do_indent(stream, indent);
    fprintf(stream, "Rows (%d):\n", lxs->n_rows);

  indent += 2;
  for (i = 0; i < lxs->n_rows; i++)
    libo_xl_row_dump(lxs->row[i], stream, indent);

  return;
}

  /**
   *  @fn libo_xl_sheet *libo_xl_sheet_meta_read(xmlDocPtr doc, int n)
   *
   *  @brief creates new @a libo_xl_sheet struct from work sheet at position @p n in file in @a doc
   *
   *  @param doc - pointer to existing XML document
   *  @param n - index of work sheet to extract
   *
   *  @return pointer to new @a libo_xl_sheet struct
   */

libo_xl_sheet *libo_xl_sheet_meta_read(xmlDocPtr doc, int n)
{
  libo_xl_sheet *sheet = NULL;
  xmlXPathContextPtr xpathCtx; 
  xmlXPathObjectPtr xpathObj; 
  xmlChar *xpathExpr = (xmlChar *)"/*[local-name() = 'workbook']/*[local-name() = 'sheets']/*[local-name() = 'sheet']";
  xmlNodeSetPtr nodes;
  xmlNodePtr node;
  int i;

  if (!doc) return NULL;
  if (n < 0) return NULL;

  sheet = libo_xl_sheet_new();
  if (!sheet) return NULL;

  xpathCtx = xmlXPathNewContext(doc);
  if(!xpathCtx)
  {
    fprintf(stderr,"Error: unable to create new XPath context\n"); fflush(stdout);
    return NULL;
  }

  xpathObj = xmlXPathEvalExpression(xpathExpr, xpathCtx);
  if(!xpathObj)
  {
      fprintf(stderr,"Error: unable to evaluate xpath expression \"%s\"\n", xpathExpr);
      xmlXPathFreeContext(xpathCtx); 
      return NULL;
  }

  nodes = xpathObj->nodesetval;
  if (nodes)
  {
    for (i = 0; i < nodes->nodeNr; i++)
    {
      if (!nodes->nodeTab[i]) continue;
      if (nodes->nodeTab[i]->type == XML_ELEMENT_NODE)
      {
        node = nodes->nodeTab[i];
        if (i == n)
        {
          sheet->name = strdup((char *)xmlGetProp(node, (xmlChar *)"name"));
          sheet->ID = atoi((char *)xmlGetProp(node, (xmlChar *)"sheetId"));
          sheet->rID = strdup((char *)xmlGetProp(node, (xmlChar *)"id"));
          break;
        }
      }
    }
  }

  xmlXPathFreeObject(xpathObj);
  xmlXPathFreeContext(xpathCtx); 

  return sheet;
}

  /**
   *  @fn void libo_xl_sheet_read(libo *l, libo_xl_sheet *sheet, int n)
   *
   *  @brief builds data for sheet number @p n into @p sheet that resides in @p l
   *
   *  @param l - pointer to existing @a libo struct
   *  @param sheet = pointer to existing @a libo_xl_sheet, filled with meta data
   *  @param n - index of work sheet to extract
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_sheet_read(libo *l, libo_xl_sheet *sheet, int n)
{
  char path[4096];
  zip_file_t *zf;
  int len;
  char *buf;
  zip_stat_t stat;
  xmlDocPtr doc = NULL;

  if (!l) return;
  if (!sheet) return;
  if (n < 0) return;

  memset(path, 0, 4096);

  sprintf(path, "xl/worksheets/sheet%d.xml", n+1);

    // open xl/worksheets/sheetN.xml

  if (zip_stat(l->z, path, 0, &stat)) return;
  if (!((stat.valid & ZIP_STAT_NAME) && (stat.valid & ZIP_STAT_SIZE))) return;

  if (strcmp(path, stat.name)) return;

  len = stat.size;

  zf = zip_fopen(l->z, path, 0);
  if (!zf)
  {
    fprintf(stderr, "Can not open '%s'\n", path); fflush(stderr);
    return;
  }

  buf = (char *)malloc(len+1);
  if (!buf)
  {
    zip_fclose(zf);
    return;
  }
  memset(buf, 0, len+1);

  zip_fread(zf, buf, len);

  zip_fclose(zf);

  doc = xmlParseMemory(buf, len);
  if (!doc)
  {
    fprintf(stderr, "Failed to parse '%s'\n", path); fflush(stderr);
    return;
  }
  free(buf);

  sheet->n_rows = count_sheet_rows_in_xml(doc);
  sheet->n_cols = count_sheet_columns_in_xml(doc);
  sheet->row = libo_xl_sheet_rows_read(sheet, doc);

  xmlFreeDoc(doc);

  return;
}

  /**
   *  @fn libo_xl_row **libo_xl_sheet_rows_read(libo_xl_sheet *sheet, xmlDocPtr doc)
   *
   *  @brief builds rows for sheet number @p sheet from @p doc
   *
   *  @param sheet = pointer to existing @a libo_xl_sheet, filled with meta data
   *  @param doc - pointer to existing XML document
   *
   *  @return pointer to new @a libo_xl_sheet struct
   */

libo_xl_row **libo_xl_sheet_rows_read(libo_xl_sheet *sheet, xmlDocPtr doc)
{
  int i, j;
  libo_xl_row **rows = NULL;
  libo_xl_row *row;
  libo_xl_cell *cell;
  xmlXPathContextPtr xpathCtx; 
  xmlXPathObjectPtr xpathObj; 
  xmlChar *xpathExpr = (xmlChar *)"/*[local-name() = 'worksheet']/*[local-name() = 'sheetData']/*[local-name() = 'row']";
  xmlNodeSetPtr nodes;
  xmlNodePtr node;
  xmlNodePtr node2;
  xmlNodePtr node3;
  int k;
  int r,c;
  char *ref;

  if (!sheet) return NULL;
  if (!doc) return NULL;

  rows = (libo_xl_row **)malloc(sizeof(libo_xl_row *) * sheet->n_rows);
  if (!rows) return NULL;
  memset(rows, 0, sizeof(libo_xl_row *) * sheet->n_rows);

  for (i = 0; i < sheet->n_rows; i++)
  {
    row = rows[i] = libo_xl_row_new();

    row->cell = (libo_xl_cell **)malloc(sizeof(libo_xl_cell *) * sheet->n_cols);
    if (!row->cell) break;
    memset(row->cell, 0, sizeof(libo_xl_cell *) * sheet->n_cols);
    row->n_cells = sheet->n_cols;

    for (j = 0; j < sheet->n_cols; j++)
      row->cell[j] = libo_xl_cell_new();
  }

  xpathCtx = xmlXPathNewContext(doc);
  if(!xpathCtx)
  {
    fprintf(stderr,"Error: unable to create new XPath context\n"); fflush(stdout);
    return 0;
  }

  xpathObj = xmlXPathEvalExpression(xpathExpr, xpathCtx);
  if(!xpathObj)
  {
      fprintf(stderr,"Error: unable to evaluate xpath expression \"%s\"\n", xpathExpr);
      xmlXPathFreeContext(xpathCtx); 
      return 0;
  }

  nodes = xpathObj->nodesetval;
  if (nodes)
  {
    if (nodes->nodeNr != sheet->n_rows)
      fprintf(stderr, "Row count does not match XML nodes\n");
    else
    {
      for (i = 0; i < sheet->n_rows; i++)
      {
        row = rows[i];

        if (!nodes->nodeTab[i]) continue;
        node = nodes->nodeTab[i];
        if (node->type == XML_ELEMENT_NODE)
        {
          if (!strcmp((char *)node->name, "row"))
          {
            node2 = node->xmlChildrenNode;

            for (j = 0; node2 && (j < sheet->n_cols);)
            {
              if (node->type == XML_ELEMENT_NODE)
              {
                if (!strcmp((char *)node2->name, "c"))
                {
                  ref = (char *)xmlGetProp(node2, (xmlChar *)"r");
                  cell_ref_to_row_col(ref, &r, &c);

                  while (j < c)
                  {
                    cell = row->cell[j];
                    cell->type = libo_xl_cell_type_expression;
                    cell->expression.value = strdup((char *)"");
                    ++j;
                  }

                  cell = row->cell[j];

                  if (xmlGetProp(node2, (xmlChar *)"t"))
                    cell->type = string_to_libo_xl_cell_type((char *)xmlGetProp(node2, (xmlChar *)"t"));
                  else
                    cell->type = libo_xl_cell_type_number;

                  switch (cell->type)
                  {
                    case libo_xl_cell_type_none:
                      break;
                    case libo_xl_cell_type_reference:
                      for (node3 = node2->xmlChildrenNode; node3; node3 = node3->next)
                      {
                        if (node3->type == XML_ELEMENT_NODE)
                        {
                          if (!strcmp((char *)node3->name, "v"))
                            cell->reference = atoi((char *)xmlNodeGetContent(node3));
                        }
                      }
                      break;
                    case libo_xl_cell_type_expression:
                      for (node3 = node2->xmlChildrenNode; node3; node3 = node3->next)
                      {
                        if (node3->type == XML_ELEMENT_NODE)
                        {
                          if (!strcmp((char *)node3->name, "f"))
                            cell->expression.formula = strdup((char *)xmlNodeGetContent(node3));
                          else if (!strcmp((char *)node3->name, "v"))
                            cell->expression.value = strdup((char *)xmlNodeGetContent(node3));
                        }
                      }
                      break;
                    case libo_xl_cell_type_number:
                      for (node3 = node2->xmlChildrenNode; node3; node3 = node3->next)
                      {
                        if (node3->type == XML_ELEMENT_NODE)
                        {
                          if (!strcmp((char *)node3->name, "v"))
                            cell->number = atof((char *)xmlNodeGetContent(node3));
                        }
                      }
                      break;
                  }
                }
                ++k;
                ++j;
              }
              node2 = node2->next;
            }

            while (j < sheet->n_cols)
            {
              cell = row->cell[j];
              cell->type = libo_xl_cell_type_expression;
              cell->expression.value = strdup((char *)"");
              ++j;
            }

          }
        }
      }
    }
  }

  xmlXPathFreeObject(xpathObj);
  xmlXPathFreeContext(xpathCtx); 

  return rows;
}

  /**
   *  @fn void libo_xl_row_dump(libo_xl_row *row, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p row to @p stream, default is STDOUT
   *
   *  @param row - pointer to existing @a libo_xl_row struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_row_dump(libo_xl_row *row, FILE *stream, int indent)
{
  int i;

  if (!row) return;

  do_indent(stream, indent); fprintf(stream, "Row:\n");
  indent += 2;

  do_indent(stream, indent);
    fprintf(stream, "Cells (%d):\n", row->n_cells);
  indent += 2;

  for (i = 0; i < row->n_cells; i++)
    libo_xl_cell_dump(row->cell[i], stream, indent);

  return;
}

  /**
   *  @fn void libo_xl_cell_dump(libo_xl_cell *cell, FILE *stream, int indent)
   *
   *  @brief dumps contents of @p cell to @p stream, default is STDOUT
   *
   *  @param cell - pointer to existing @a libo_xl_cell struct
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to indent output
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_cell_dump(libo_xl_cell *cell, FILE *stream, int indent)
{
  string *str;

  if (!cell) return;
  if (!stream) stream = stdout;

  do_indent(stream, indent); fprintf(stream, "Cell:\n");
  indent += 2;

  do_indent(stream, indent);
    fprintf(stream, "Type: %s\n",
            libo_xl_cell_type_to_string(cell->type));

  do_indent(stream, indent); fprintf(stream, "Contents:\n");
  indent += 2;

  switch (cell->type)
  {
    case libo_xl_cell_type_none:
      do_indent(stream, indent); fprintf(stream, "[NONE]\n");
      break;
    case libo_xl_cell_type_reference:
      do_indent(stream, indent);
        fprintf(stream, "Reference: %d\n", cell->reference);
      do_indent(stream, indent);
        str = strings_find_by_id(__xl__->strings, cell->reference);
        fprintf(stream, "String: %s\n", str ? str->text : "");
      break;
    case libo_xl_cell_type_expression:
      do_indent(stream, indent); fprintf(stream, "Expression:\n");
      indent += 2;
      do_indent(stream, indent);
        fprintf(stream, "Formula: %s\n", cell->expression.formula);
      do_indent(stream, indent);
        fprintf(stream, "Value: %s\n", cell->expression.value);
      break;
    case libo_xl_cell_type_number:
      do_indent(stream, indent);
        fprintf(stream, "Number: %f\n", cell->number);
      break;
  }

  return;
}

  /**
   *  @fn char *libo_xl_cell_type_to_string(libo_xl_cell_type ct)
   *
   *  @brief returns string representation of @p ct
   *
   *  @param ct - @a libo_xl_cell_type
   *
   *  @return string representation of @p ct
   */

char *libo_xl_cell_type_to_string(libo_xl_cell_type ct)
{
  switch (ct)
  {
    case libo_xl_cell_type_none: return "[UNKNOWN]";
    case libo_xl_cell_type_reference: return "REFERENCE";
    case libo_xl_cell_type_expression: return "EXPRESSION";
    case libo_xl_cell_type_number: return "NUMBER";
  }

  return "[UNKNOWN]";
}

  /**
   *  @fn libo_xl_column *libo_xl_column_new(void)
   *
   *  @brief creates a new @a libo_xl_column struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo_xl_column struct
   */

libo_xl_column *libo_xl_column_new(void)
{
  libo_xl_column *c = malloc(sizeof(libo_xl_column));
  if (c) memset(c, 0, sizeof(libo_xl_column));

  return c;
}

  /**
   *  @fn libo_xl_column *libo_xl_column_new_with_values(float width, int autowidth)
   *
   *  @brief creates a new @a libo_xl_column struct with values filled
   *
   *  @param width - specific width of column
   *  @param autowidth - flag, 1 to auto-caculate column width, 0 to use specific width
   *
   *  @return pointer to new @a libo_xl_column struct
   */

libo_xl_column *libo_xl_column_new_with_values(float width, int autowidth)
{
  libo_xl_column *c = libo_xl_column_new();
  if (c)
  {
    c->width = width;
    c->autowidth = autowidth;
  }

  return c;
}

  /**
   *  @fn void libo_xl_column_free(libo_xl_column *column)
   *
   *  @brief frees all memory allocated to @p column
   *
   *  @param column - pointer to existing @a libo_xl_column struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_column_free(libo_xl_column *column) { if (column) free(column); }

  /**
   *  @fn libo_xl_filter *libo_xl_filter_new(void)
   *
   *  @brief returns new @a libo_xl_filter struct
   *
   *  @par Parameters
   *  None.
   *
   *  @return pointer to new @a libo_xl_filter struct
   */

libo_xl_filter *libo_xl_filter_new(void)
{
  libo_xl_filter *filter;

  filter = (libo_xl_filter *)malloc(sizeof(libo_xl_filter));
  if (filter) memset(filter, 0, sizeof(libo_xl_filter));

  return filter;
}

  /**
   *  @fn libo_xl_filter *libo_xl_filter_new_with_values(
   *                         unsigned int first_column,
   *                         unsigned int last_column)
   *
   *  @brief returns new @a libo_xl_filter struct
   *
   *  @param first_column - first column that defines filter
   *  @param last_column - last column that defines filter
   *
   *  @return pointer to new @a libo_xl_filter struct
   */

libo_xl_filter *libo_xl_filter_new_with_values(unsigned int first_column,
                                               unsigned int last_column)
{
  libo_xl_filter *filter;

  filter = libo_xl_filter_new();
  if (filter)
  {
    filter->first_column = first_column;
    filter->last_column = last_column;
  }

  return filter;
}

  /**
   *  @fn void libo_xl_filter_free(libo_xl_filter *filter)
   *
   *  @brief frees all memory allocated to @p filter
   *
   *  @param filter - pointer to existing @a libo_xl_filter struct
   *
   *  @par Returns
   *  Nothing.
   */

void libo_xl_filter_free(libo_xl_filter *filter)
{
  if (!filter) return;

  free(filter);
}

 // INTERNALS

  /**
   *  @fn static void cell_ref_to_row_col(char *ref, int *row, int *col)
   *
   *  @brief converts string cell reference to row and column
   *
   *  @p ref is in form "A1", or "YX120"
   *
   *  @param ref - string cell reference
   *  @param row - pointer to integer to store extracted row
   *  @param col - pointer to integer to store extracted col
   *
   *  @par Returns
   *  Nothing.
   */

static void cell_ref_to_row_col(char *ref, int *row, int *col)
{
  char *p;
  int c;
  int i;
  int prod;
  int mag = 0;

  *row = *col = 0;

  if (!ref || !row || !col) return;

  p = ref;
  while (isalpha(*p)) ++p;
  *row = atoi(p) - 1;

  --p;

  while (p >= ref)
  {
    c = (*p - 'A' + 1);
    for (prod = 1, i = 0; i < mag; i++)
      prod *= 26;
    *col += c * prod;
    ++mag;
    --p;
  }

  *col -= 1;

  return;
}

  /**
   *  @fn static int is_office(libo *l)
   *
   *  @brief determines if contents of @p l is legitimate Office document
   *
   *  @param l - pointer to existing @a libo struct, with ZIP contents
   *
   *  @return 1 if Office document, 0 for anything else
   */

static int is_office(libo *l)
{
  zip_int64_t n_entries;
  int i;
  char *name;
  int have_core = 0;
  int have_app = 0;

  if (!l) return 0;
  if (!l->z) return 0;

  n_entries = zip_get_num_entries(l->z, 0);

  for (i = 0; i < n_entries; i++)
  {
    name = (char *)zip_get_name(l->z, i, 0);
    if (!strcmp(name, "docProps/core.xml")) have_core = 1;
    if (!strcmp(name, "docProps/app.xml")) have_app = 1;

    if (have_core && have_app) return 1;
  }

  return 0;
}

  /**
   *  @fn static void do_indent(FILE *stream, int indent)
   *
   *  @brief emits @p indent spaces to @p stream
   *
   *  @param stream - open FILE * for writing
   *  @param indent - number of spaces to output
   *
   *  @par Returns
   *  Nothing.
   */

static void do_indent(FILE *stream, int indent)
{
  int i;

  if (!stream) return;

  for (i = 0; i < indent; i++) fputc(' ', stream);

  return;
}

  /**
   *  @fn static int is_supported(libo *l)
   *
   *  @brief determines if this Office document is currently implemented
   *
   *  @param l - pointer to existing @a libo struct, with ZIP contents
   *
   *  @return 1 implemented, 0 otherwise
   */

static int is_supported(libo *l)
{
  if (!l) return 0;

  switch (l->type)
  {
    case libo_type_xl:
      return 1;

    case libo_type_none:
    case libo_type_doc:
    case libo_type_pp:
    default:
      return 0;
  }

  return 0;
}

  /**
   *  @fn static libo_type get_type(libo *l)
   *
   *  @brief returns Office document type from @p l
   *
   *  @param l - pointer to existing @a libo struct, with ZIP contents
   *
   *  @return @a libo_type of document
   */

static libo_type get_type(libo *l)
{
  zip_file_t *zf;
  int len;
  char *buf;
  char *app_file_name = "docProps/app.xml";
  zip_stat_t stat;
  xmlDocPtr doc = NULL;
  libo_type type = libo_type_none;

  if (!l) return libo_type_none;

    // open docProps/app.xml

  if (zip_stat(l->z, app_file_name, 0, &stat)) return libo_type_none;
  if (!((stat.valid & ZIP_STAT_NAME) && (stat.valid & ZIP_STAT_SIZE))) return libo_type_none;

  if (strcmp(app_file_name, stat.name)) return libo_type_none;

  len = stat.size;

  zf = zip_fopen(l->z, app_file_name, 0);
  if (!zf)
  {
    fprintf(stderr, "Can not open '%s'\n", app_file_name); fflush(stderr);
    return libo_type_none;
  }

  buf = (char *)malloc(len+1);
  if (!buf)
  {
    zip_fclose(zf);
    return libo_type_none;
  }
  memset(buf, 0, len+1);

  zip_fread(zf, buf, len);

  zip_fclose(zf);

  doc = xmlParseMemory(buf, len);
  if (!doc)
  {
    fprintf(stderr, "Failed to parse app.xml\n"); fflush(stderr);
    return libo_type_none;
  }
  free(buf);

    // find entry for document type

  buf = find_app_type_name_in_xml(doc);

  xmlFreeDoc(doc);

  if (!buf) return libo_type_none;

    // convert to libo_type

  if (!strcmp(buf, "Microsoft Excel")) type = libo_type_xl;

  free(buf);

  return type;
}

  /**
   *  @fn static char *find_app_type_name_in_xml(xmlDocPtr doc)
   *
   *  @brief locates application type name in XML document
   *
   *  @param doc - pointer to XML document
   *
   *  @return string containing application type name
   */

static char *find_app_type_name_in_xml(xmlDocPtr doc)
{
  xmlXPathContextPtr xpathCtx; 
  xmlXPathObjectPtr xpathObj; 
  xmlChar *xpathExpr = (xmlChar *)"/*[local-name() = 'Properties']/*[local-name() = 'Application']";
  xmlNodeSetPtr nodes;
  xmlNodePtr node;
  char *app_type_name = NULL;

  if (!doc) return NULL;

  xpathCtx = xmlXPathNewContext(doc);
  if(!xpathCtx)
  {
    fprintf(stderr,"Error: unable to create new XPath context\n"); fflush(stdout);
    return NULL;
  }

  xpathObj = xmlXPathEvalExpression(xpathExpr, xpathCtx);
  if(!xpathObj)
  {
      fprintf(stderr,"Error: unable to evaluate xpath expression \"%s\"\n", xpathExpr);
      xmlXPathFreeContext(xpathCtx); 
      return NULL;
  }

  nodes = xpathObj->nodesetval;
  if (nodes)
  {
    for (int i = 0; i < nodes->nodeNr; i++)
    {
      if (!nodes->nodeTab[i]) continue;
      if (nodes->nodeTab[i]->type == XML_ELEMENT_NODE)
      {
        node = nodes->nodeTab[i];
        app_type_name = (char *)xmlNodeListGetString(doc, node->xmlChildrenNode, 1);
        break;
      }
    }
  }

  xmlXPathFreeObject(xpathObj);
  xmlXPathFreeContext(xpathCtx); 

  return app_type_name;
}

  /**
   *  @fn static int count_sheets_in_xml(xmlDocPtr doc)
   *
   *  @brief returns number of work sheets in XML document
   *
   *  @param doc - pointer to XML document
   *
   *  @return count of work sheets
   */

static int count_sheets_in_xml(xmlDocPtr doc)
{
  xmlXPathContextPtr xpathCtx; 
  xmlXPathObjectPtr xpathObj; 
  xmlChar *xpathExpr = (xmlChar *)"/*[local-name() = 'workbook']/*[local-name() = 'sheets']";
  xmlNodeSetPtr nodes;
  xmlNodePtr node;
  int count = 0;

  if (!doc) return 0;

  xpathCtx = xmlXPathNewContext(doc);
  if(!xpathCtx)
  {
    fprintf(stderr,"Error: unable to create new XPath context\n"); fflush(stdout);
    return 0;
  }

  xpathObj = xmlXPathEvalExpression(xpathExpr, xpathCtx);
  if(!xpathObj)
  {
      fprintf(stderr,"Error: unable to evaluate xpath expression \"%s\"\n", xpathExpr);
      xmlXPathFreeContext(xpathCtx); 
      return 0;
  }

  nodes = xpathObj->nodesetval;
  if (nodes)
  {
    for (int i = 0; i < nodes->nodeNr; i++)
    {
      if (!nodes->nodeTab[i]) continue;
      if (nodes->nodeTab[i]->type == XML_ELEMENT_NODE)
      {
        node = nodes->nodeTab[i];
        if (!strcmp((char *)node->name, "sheets"))
        {
          count = xmlChildElementCount(node);
          break;
        }
      }
    }
  }

  xmlXPathFreeObject(xpathObj);
  xmlXPathFreeContext(xpathCtx); 

  return count;
}

  /**
   *  @fn static int count_sheet_rows_in_xml(xmlDocPtr doc)
   *
   *  @brief returns number of work sheet rows in XML document
   *
   *  @param doc - pointer to XML document
   *
   *  @return count of rows
   */

static int count_sheet_rows_in_xml(xmlDocPtr doc)
{
  xmlXPathContextPtr xpathCtx; 
  xmlXPathObjectPtr xpathObj; 
  xmlChar *xpathExpr = (xmlChar *)"/*[local-name() = 'worksheet']/*[local-name() = 'sheetData']";
  xmlNodeSetPtr nodes;
  xmlNodePtr node;
  int count = 0;

  if (!doc) return 0;

  xpathCtx = xmlXPathNewContext(doc);
  if(!xpathCtx)
  {
    fprintf(stderr,"Error: unable to create new XPath context\n"); fflush(stdout);
    return 0;
  }

  xpathObj = xmlXPathEvalExpression(xpathExpr, xpathCtx);
  if(!xpathObj)
  {
      fprintf(stderr,"Error: unable to evaluate xpath expression \"%s\"\n", xpathExpr);
      xmlXPathFreeContext(xpathCtx); 
      return 0;
  }

  nodes = xpathObj->nodesetval;
  if (nodes)
  {
    for (int i = 0; i < nodes->nodeNr; i++)
    {
      if (!nodes->nodeTab[i]) continue;
      if (nodes->nodeTab[i]->type == XML_ELEMENT_NODE)
      {
        node = nodes->nodeTab[i];
        if (!strcmp((char *)node->name, "sheetData"))
        {
          count = xmlChildElementCount(node);
          break;
        }
      }
    }
  }

  xmlXPathFreeObject(xpathObj);
  xmlXPathFreeContext(xpathCtx); 

  return count;
}

  /**
   *  @fn static int count_sheet_columns_in_xml(xmlDocPtr doc)
   *
   *  @brief returns number of work sheet columns in XML document
   *
   *  @param doc - pointer to XML document
   *
   *  @return count of columns
   */

static int count_sheet_columns_in_xml(xmlDocPtr doc)
{
  xmlXPathContextPtr xpathCtx; 
  xmlXPathObjectPtr xpathObj; 
  xmlChar *xpathExpr = (xmlChar *)"/*[local-name() = 'worksheet']/*[local-name() = 'sheetData']/*[local-name() = 'row']";
  xmlNodeSetPtr nodes;
  xmlNodePtr node;
  int count = 0;
  char *spans;
  char *p;

  if (!doc) return 0;

  xpathCtx = xmlXPathNewContext(doc);
  if(!xpathCtx)
  {
    fprintf(stderr,"Error: unable to create new XPath context\n"); fflush(stdout);
    return 0;
  }

  xpathObj = xmlXPathEvalExpression(xpathExpr, xpathCtx);
  if(!xpathObj)
  {
      fprintf(stderr,"Error: unable to evaluate xpath expression \"%s\"\n", xpathExpr);
      xmlXPathFreeContext(xpathCtx); 
      return 0;
  }

  nodes = xpathObj->nodesetval;
  if (nodes)
  {
    for (int i = 0; i < nodes->nodeNr; i++)
    {
      if (!nodes->nodeTab[i]) continue;
      if (nodes->nodeTab[i]->type == XML_ELEMENT_NODE)
      {
        node = nodes->nodeTab[i];
        if (!strcmp((char *)node->name, "row"))
        {
          spans = (char *)xmlGetProp(node, (xmlChar *)"spans");
          if (spans)
          {
            p = strchr(spans, ':');
            if (p)
            {
              ++p;
              count = atoi(p);
            }
          }
          break;
        }
      }
    }
  }

  xmlXPathFreeObject(xpathObj);
  xmlXPathFreeContext(xpathCtx); 

  return count;
}

  /**
   *  @fn static libo_xl_cell_type string_to_libo_xl_cell_type(char *s)
   *
   *  @brief converts string representation of cell type to @a libo_xl_cell_type
   *
   *  @param s - string representation of cell type
   *
   *  @return @a libo_xl_cell_type
   */

static libo_xl_cell_type string_to_libo_xl_cell_type(char *s)
{
  if (!s) return libo_xl_cell_type_none;

  if (!strcmp(s, "s")) return libo_xl_cell_type_reference;
  if (!strcmp(s, "e")) return libo_xl_cell_type_expression;

  return libo_xl_cell_type_none;
}

  /**
   *  @fn static int libo_xl_write(libo *l)
   *
   *  @brief writes XL document to file
   *
   *  @param l - pointer to existing @a libo struct
   *
   *  @return 0 on success, STDIO error on failure
   */

static int libo_xl_write(libo *l)
{
  int success = -1;

    // create required directories in zip file

  if (!l) goto bail;
  if (!l->z) goto bail;

/*
  if ((zip_dir_add(l->z, "_rels", 0)) < 0) goto bail;
  if ((zip_dir_add(l->z, "docProps", 0)) < 0) goto bail;
  if ((zip_dir_add(l->z, "xl", 0)) < 0) goto bail;
  if ((zip_dir_add(l->z, "xl/_rels", 0)) < 0) goto bail;
  if ((zip_dir_add(l->z, "xl/theme", 0)) < 0) goto bail;
  if ((zip_dir_add(l->z, "xl/worksheets", 0)) < 0) goto bail;
*/

  libo_xl_content_types_write(l);
  libo_xl_docprops_write(l);
  libo_xl__rels_dot_rels_write(l);
  libo_xl_xl_rels_write(l);
  libo_xl_themes_write(l);
  libo_xl_styles_write(l);
  libo_xl_workbook_write(l);
  libo_xl_sheets_write(l);
  libo_xl_shared_strings_write(l);

  success = 0;

bail:
  return success;
}

#include "libo-xl-theme.c"

  /**
   *  @fn int libo_xl_themes_write(libo *l)
   *
   *  @brief writes XL themes to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_themes_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;

  if (!l || !l->z) return -1;

  zs = zip_source_buffer_create(libo_xl_theme_standard, strlen(libo_xl_theme_standard), 0, &err);
  if (!zs) return -1;

  if ((zip_file_add(l->z, "xl/theme/theme1.xml", zs, 0)) < 0) return -1;

  return 0;
}

#include "libo-xl-styles.c"

  /**
   *  @fn int libo_xl_styles_write(libo *l)
   *
   *  @brief writes XL styles to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_styles_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;

  if (!l || !l->z) return -1;

  zs = zip_source_buffer_create(libo_xl_styles_standard, strlen(libo_xl_styles_standard), 0, &err);
  if (!zs) return -1;

  if ((zip_file_add(l->z, "xl/styles.xml", zs, 0)) < 0) return -1;

  return 0;
}

  /**
   *  @fn int libo_xl_docprops_write(libo *l)
   *
   *  @brief writes XL document properties to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_docprops_write(libo *l)
{
  int success = -1;

  if (!l || !l->z) goto bail;

  if (libo_xl_docprops_app_write(l)) goto bail;
  if (libo_xl_docprops_core_write(l)) goto bail;

  success = 0;

bail:
  return success;
}

static char *libo_xl_app_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
  "<Application>Microsoft Excel</Application>"
  "<DocSecurity>0</DocSecurity>"
  "<ScaleCrop>false</ScaleCrop>";

static char *libo_xl_app_boiler_plate_2 =  /**<  XML boiler plate  */
  "<LinksUpToDate>false</LinksUpToDate>"
  "<SharedDoc>false</SharedDoc>"
  "<HyperlinksChanged>false</HyperlinksChanged>"
  "<AppVersion>16.0300</AppVersion>"
  "</Properties>";

  /**
   *  @fn int libo_xl_docprops_app_write(libo *l)
   *
   *  @brief writes application document properties to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_docprops_app_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  int success = -1;
  char number[25];
  char *buf = NULL;
  int i;

  if (!l || !l->z) goto bail;

  memset(number, 0, 25);

  buf = strapp(buf, libo_xl_app_boiler_plate_1);

  buf = strapp(buf, "<HeadingPairs>");
  buf = strapp(buf, "<vt:vector size=\"2\" baseType=\"variant\">");
  buf = strapp(buf, "<vt:variant>");
  buf = strapp(buf, "<vt:lpstr>Worksheets</vt:lpstr>");
  buf = strapp(buf, "</vt:variant>");
  buf = strapp(buf, "<vt:variant>");
  buf = strapp(buf, "<vt:i4>");
  sprintf(number, "%d", l->xl->book->n_sheets);
  buf = strapp(buf, number);
  buf = strapp(buf, "</vt:i4>");
  buf = strapp(buf, "</vt:variant>");
  buf = strapp(buf, "</vt:vector>");
  buf = strapp(buf, "</HeadingPairs>");
  buf = strapp(buf, "<TitlesOfParts>");
  buf = strapp(buf, "<vt:vector size=\"");
  sprintf(number, "%d", l->xl->book->n_sheets);
  buf = strapp(buf, number);
  buf = strapp(buf, "\" baseType=\"lpstr\">");
  for (i = 0; i < l->xl->book->n_sheets; i++)
  {
    buf = strapp(buf, "<vt:lpstr>");
    buf = strapp(buf, l->xl->book->sheet[i]->name);
    buf = strapp(buf, "</vt:lpstr>");
  }
  buf = strapp(buf, "</vt:vector>");
  buf = strapp(buf, "</TitlesOfParts>");

  buf = strapp(buf, libo_xl_app_boiler_plate_2);

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  if ((zip_file_add(l->z, "docProps/app.xml", zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

static char *libo_xl_core_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
  "<dc:creator>LIBO</dc:creator>"
  "<cp:lastModifiedBy>LIBO</cp:lastModifiedBy>";

static char *libo_xl_core_boiler_plate_2 =  /**<  XML boiler plate  */
  "</cp:coreProperties>";

  /**
   *  @fn int libo_xl_docprops_core_write(libo *l)
   *
   *  @brief writes XL core document properties to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_docprops_core_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  int success = -1;
  char *buf = NULL;
  char date[50];
  time_t tim = 0;
  struct tm *gmt;

  if (!l || !l->z) goto bail;

  memset(date, 0, 50);

  tim = time(NULL);
  gmt = gmtime(&tim);
  strftime(date, 50, "%Y-%m-%dT%H:%M:%SZ", gmt);

  buf = strapp(buf, libo_xl_core_boiler_plate_1);

  buf = strapp(buf, "<dcterms:created xsi:type=\"dcterms:W3CDTF\">");
  buf = strapp(buf, date);
  buf = strapp(buf, "</dcterms:created>");
  buf = strapp(buf, "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">");
  buf = strapp(buf, date);
  buf = strapp(buf, "</dcterms:modified>");

  buf = strapp(buf, libo_xl_core_boiler_plate_2);

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  if ((zip_file_add(l->z, "docProps/core.xml", zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

  /**
   *  @fn int libo_xl_xl_rels_write(libo *l)
   *
   *  @brief writes XL relationships to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_xl_rels_write(libo *l)
{
  int success = -1;

  if (!l || !l->z) goto bail;

  if (libo_xl_xl_rels_workbook_rels_write(l)) goto bail;

  success = 0;

bail:
  return success;
}

static char *libo_xl__rels_dot_rels_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
  "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
  "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
  "</Relationships>";

  /**
   *  @fn int libo_xl__rels_dot_rels_write(libo *l)
   *
   *  @brief writes XL default relationships to document files
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl__rels_dot_rels_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  int success = -1;
  char *buf = NULL;

  buf = strapp(buf, libo_xl__rels_dot_rels_boiler_plate_1);

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  if ((zip_file_add(l->z, "_rels/.rels", zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

static char *libo_xl_workbook_rels_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
  "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
  "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";

static char *libo_xl_workbook_rels_boiler_plate_2 =  /**<  XML boiler plate  */
  "</Relationships>";

  /**
   *  @fn int libo_xl_xl_rels_workbook_rels_write(libo *l)
   *
   *  @brief writes XL workbook relationships to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_xl_rels_workbook_rels_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  int success = -1;
  char *buf = NULL;
  char number[25];
  int i;

  if (!l || !l->z) goto bail;

  memset(number, 0, 25);

  buf = strapp(buf, libo_xl_workbook_rels_boiler_plate_1);

  for (i = 0; i < l->xl->book->n_sheets; i++)
  {
    sprintf(number, "%d", i+4);
    buf = strapp(buf, "<Relationship Id=\"rId");
    buf = strapp(buf, number);
    buf = strapp(buf, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
    sprintf(number, "%d", i+1);
    buf = strapp(buf, number);
    buf = strapp(buf, ".xml\"/>");
  }

  buf = strapp(buf, libo_xl_workbook_rels_boiler_plate_2);

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  if ((zip_file_add(l->z, "xl/_rels/workbook.xml.rels", zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

static char *libo_xl_content_types_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
  //"<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>\n"
  "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
  "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n";

static char *libo_xl_content_types_boiler_plate_2 =  /**<  XML boiler plate  */
  "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n"
  "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n"
  "</Types>\n";

  /**
   *  @fn int libo_xl_content_types_write(libo *l)
   *
   *  @brief writes XL content types to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_content_types_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  char *buf = NULL;
  char number[25];
  int i;
  int success = -1;

  if (!l || !l->z) goto bail;

  memset(number, 0, 25);

  buf = strapp(buf, libo_xl_content_types_boiler_plate_1);

  switch (l->type)
  {
    case libo_type_xl:
      buf = strapp(buf, "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");

      for (i = 0; i < l->xl->book->n_sheets; i++)
      {
        buf = strapp(buf, "<Override PartName=\"/xl/worksheets/sheet");
        sprintf(number, "%d", i+1);
        buf = strapp(buf, number);
        buf = strapp(buf, ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
      }

      buf = strapp(buf, "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
      buf = strapp(buf, "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
      buf = strapp(buf, "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");

      break; //libo_type_xl

    case libo_type_doc:
      break; //libo_type_doc

    case libo_type_pp:
      break; //libo_type_pp

    default: break;
  }

  buf = strapp(buf, libo_xl_content_types_boiler_plate_2);

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  if ((zip_file_add(l->z, "[Content_Types].xml", zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

static char *libo_xl_workbook_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" mc:Ignorable=\"x15 xr xr6 xr10\">";

  /**
   *  @fn int libo_xl_workbook_write(libo *l)
   *
   *  @brief writes XL workbook to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_workbook_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  char *buf = NULL;
  char number[25];
  int success = -1;
  int i;

  if (!l || !l->z) goto bail;

  buf = strapp(buf, libo_xl_workbook_boiler_plate_1);
  buf = strapp(buf, "<fileVersion appName=\"xl\" lastEdited=\"1\" lowestEdited=\"1\" rupBuild=\"25601\"/>"); // what are parameters??
  buf = strapp(buf, "<workbookPr defaultThemeVersion=\"166925\"/>");
  buf = strapp(buf, "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
  buf = strapp(buf, "<mc:Choice Requires=\"x15\">");
  buf = strapp(buf, "<x15ac:absPath xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\" url=\"");
  buf = strapp(buf, l->path);
  buf = strapp(buf, "\"/>");
  buf = strapp(buf, "</mc:Choice>");
  buf = strapp(buf, "</mc:AlternateContent>");
  buf = strapp(buf, "<xr:revisionPtr revIDLastSave=\"0\" documentId=\"13_ncr:40009_{47680350-0BCE-45AA-9C35-94426BD8D69C}\" xr6:coauthVersionLast=\"47\" xr6:coauthVersionMax=\"47\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\"/>");
  buf = strapp(buf, "<bookViews>");
  buf = strapp(buf, "<workbookView xWindow=\"-108\" yWindow=\"-108\" windowWidth=\"23256\" windowHeight=\"12576\"/>");
  buf = strapp(buf, "</bookViews>");
  buf = strapp(buf, "<sheets>");
  for (i = 0; i < l->xl->book->n_sheets; i++)
  {
    buf = strapp(buf, "<sheet name=\"");
    buf = strapp(buf, l->xl->book->sheet[i]->name);
    buf = strapp(buf, "\" sheetId=\"");
    sprintf(number, "%d", i+1);
    buf = strapp(buf, number);
    buf = strapp(buf, "\" r:id=\"rId");
    sprintf(number, "%d", i+4);
    buf = strapp(buf, number);
    buf = strapp(buf, "\"/>");
  }
  buf = strapp(buf, "</sheets>");
  buf = strapp(buf, "<calcPr calcId=\"0\"/>");
  buf = strapp(buf, "</workbook>");

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  if ((zip_file_add(l->z, "xl/workbook.xml", zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

  /**
   *  @fn int libo_xl_sheets_write(libo *l)
   *
   *  @brief writes XL worksheets to document file
   *
   *  @param l - pointer to existing @a libo document
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_sheets_write(libo *l)
{
  int success = -1;
  int i;

  if (!l) goto bail;
  if (l->type != libo_type_xl) goto bail;

  libo_xl_renumber_strings(l);

  for (i = 0; i < l->xl->book->n_sheets; i++)
    if (libo_xl_sheet_write(l, i) < 0) goto bail;

  success = 0;

bail:
  return success;
}

  //"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">";
static char *libo_xl_sheet_boiler_plate_1 =  /**<  XML boiler plate  */
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
  "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" mc:Ignorable=\"x14ac xr xr2 xr3\" xr:uid=\"{00000000-0001-0000-0800-000000000000}\">\n";

static char *libo_xl_sheet_boiler_plate_2 =  /**<  XML boiler plate  */
  "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n"
  "<pageSetup orientation=\"portrait\" horizontalDpi=\"300\" verticalDpi=\"300\" r:id=\"rId1\"/>\n"
  "</worksheet>\n";

  /**
   *  @fn int libo_xl_sheet_write(libo *l, int sheet)
   *
   *  @brief writes XL worksheet to document file
   *
   *  @param l - pointer to existing @a libo document
   *  @param sheet - index of sheet to write 
   *
   *  @return 0 on success, -1 on failure
   */

static int libo_xl_sheet_write(libo *l, int sheet)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  char *buf = NULL;
  char number[25];
  char name[256];
  int success = -1;
  libo_xl_sheet *sht;

  if (!l) goto bail;
  if (l->type != libo_type_xl) goto bail;
  if (!l->xl) goto bail;
  if (!l->xl->book) goto bail;
  if (sheet >= l->xl->book->n_sheets) goto bail;
  if (!l->xl->book->sheet[sheet]) goto bail;

  memset(number, 0, 25);
  memset(name, 0, 256);

  sht = l->xl->book->sheet[sheet];

  libo_xl_sheet_count_columns(sht);
  if (!sht->column) sht->column = libo_xl_sheet_columns_create_defaults(sht);

  buf = strapp(buf, libo_xl_sheet_boiler_plate_1);

  libo_xl_sheet_dimension_add(l, sheet, &buf);
  libo_xl_sheet_sheetviews_add(l, sheet, &buf);
  libo_xl_sheet_formatpr_add(l, sheet, &buf);
  libo_xl_sheet_cols_add(l, sheet, &buf);
  libo_xl_sheet_sheetdata_add(l, sheet, &buf);
  libo_xl_sheet_filter_add(l, sheet, &buf);

  buf = strapp(buf, libo_xl_sheet_boiler_plate_2);

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  sprintf(name, "xl/worksheets/sheet%d.xml", sheet+1);
  if ((zip_file_add(l->z, name, zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

 /**
  * @fn static char *strapp(char *s1, char *s2)
  *
  * @brief appends @p s2 to @p s1 , reallocating memory
  *
  * NOTE:  The user is responsible for freeing @p s2 if required.
  *
  * @param s1 - original string, can be NULL
  * @param s2 - addendum to string
  *
  * @return pointer to newly formed string
  */

static char *strapp(char *s1, char *s2)
{
  int len1, len2;
  char *tmp;

  if (!s1)
  {
    s1 = malloc(1);
    if (!s1) return NULL;
    *s1 = 0;
  }

  if (!s2) return s1;

  len1 = strlen(s1);
  len2 = strlen(s2);

  tmp = realloc(s1, len1 + len2 + 1);
  if (tmp)
  {
    s1 = tmp;
    strcat(s1, s2);
  }

  return s1;
}

 /**
  * @fn static int libo_xl_shared_strings_write(libo *l)
  *
  * @brief writes XL shared strings to file
  *
  * @param l - pointer to existing @a libo struct
  *
  * @return 0 on success, STDIO error on failure
  */

static int libo_xl_shared_strings_write(libo *l)
{
  zip_error_t err;
  zip_source_t *zs = NULL;
  char *buf = NULL;
  strings *strs;
  char number[25];
  char name[256];
  int success = -1;

  if (!l) goto bail;
  if (l->type != libo_type_xl) goto bail;
  if (!l->xl) goto bail;
  if (!l->xl->strings) goto bail;

  strs = l->xl->strings;

    // Count number of strings

  _strings_count = 0;
  strings_walk(strs, string_text, libo_xl_strings_count_action);

  buf = strapp(buf, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
  sprintf(number, "%d", _strings_count);
  buf = strapp(buf, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
  buf = strapp(buf, number);
  buf = strapp(buf, "\" uniqueCount=\"");
  buf = strapp(buf, number);
  buf = strapp(buf, "\">");

    // Add each strings entry into buffer

  _strings_buf = buf;
  strings_walk(strs, string_id, libo_xl_strings_add_action);
  buf = _strings_buf;

  strapp(buf, "</sst>");

  zs = zip_source_buffer_create(buf, strlen(buf), 1, &err);
  if (!zs) goto bail;

  strcpy(name, "xl/sharedStrings.xml");
  if ((zip_file_add(l->z, name, zs, 0)) < 0) goto bail;

  success = 0;

bail:
  if (success < 0 && buf) free(buf);

  return success;
}

 /**
  * @fn static void libo_xl_renumber_strings(libo *l)
  *
  * @brief renumbers all string references in XL doc
  *
  * @param l - pointer to existing @a libo struct
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_renumber_strings(libo *l)
{
  libo_xl_book *book;
  libo_xl_sheet *sheet;
  libo_xl_row *row;
  libo_xl_cell *cell;
  int i, j, k;
  unsigned int new_id = 0;
  strings *strings = NULL;
  string *found;

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (!l->xl) return;
  if (!l->xl->strings) return;

  book = l->xl->book;
  if (!book) return;

  strings = strings_new();
  if (!strings) return;

    // Set shared string references to new strings IDs

  for (i = 0; i < book->n_sheets; i++)
  {
    sheet = book->sheet[i];
    if (!sheet) return;
    for (j = 0; j < sheet->n_rows; j++)
    {
      row = sheet->row[j];
      if (!row) return;
      for (k = 0; k < row->n_cells; k++)
      {
        cell = row->cell[k];
        switch (cell->type)
        {
          case libo_xl_cell_type_none:
          case libo_xl_cell_type_expression:
          case libo_xl_cell_type_number:
            break;

          case libo_xl_cell_type_reference:
            found = strings_find_by_id(l->xl->strings, cell->reference);
            if (found) found = strings_find_by_text(strings, found->text);
            if (found)
            {
              strings_add(strings, string_dup(found));
              cell->reference = found->id;
            }
            else
            {
              found = strings_find_by_id(l->xl->strings, cell->reference);
              if (found) strings_add(strings, string_dup(found));
              cell->reference = new_id;
              ++new_id;
            }
            break;
        }
      }
    }
  }

  strings_free(l->xl->strings);
  l->xl->strings = strings;
}

 /**
  * @fn static void libo_xl_sheet_dimension_add(libo *l,
  *                                                 int sheet,
  *                                                 char **buf)
  *
  * @brief adds worksheet dimension references to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_dimension_add(libo *l, int sheet, char **buf)
{
  char number[25];

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  memset(number, 0, 25);

    /* <dimension ref="A1:G5"/> */

  *buf = strapp(*buf, "<dimension ref=\"A1:");
  if (l->xl->book->sheet[sheet]->n_cols)
    *buf = strapp(*buf, column_number_to_reference(l->xl->book->sheet[sheet]->n_cols-1));
  else *buf = strapp(*buf, "A");
  sprintf(number, "%d", l->xl->book->sheet[sheet]->n_rows);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, "\"/>\n");
}

 /**
  * @fn static void libo_xl_sheet_sheetviews_add(libo *l,
  *                                                  int sheet,
  *                                                  char **buf)
  *
  * @brief adds XL sheetviews information to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_sheetviews_add(libo *l, int sheet, char **buf)
{
  char number[25];
  libo_xl_sheet *sht;

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  memset(number, 0, 25);

  sht = l->xl->book->sheet[sheet];

  if (!sht->freeze.type) return;

  *buf = strapp(*buf, "<sheetViews>\n");
  *buf = strapp(*buf, "<sheetView tabSelected=\"1\" topLeftCell=\"A1\" workbookViewId=\"0\">\n");
  *buf = strapp(*buf, "<pane ");
  switch (sht->freeze.type)
  {
    case libo_xl_freeze_type_none: break;
    case libo_xl_freeze_type_top:
      *buf = strapp(*buf, "ySplit");
      break;
    case libo_xl_freeze_type_left: break;
      *buf = strapp(*buf, "xSplit");
      break;
  }
  *buf = strapp(*buf, "=\"1\" topLeftCell=\"");
  switch (sht->freeze.type)
  {
    case libo_xl_freeze_type_none: break;
    case libo_xl_freeze_type_top:
      *buf = strapp(*buf, "A");
      sprintf(number, "%d", sht->freeze.n + 1);
      *buf = strapp(*buf, number);
      break;
    case libo_xl_freeze_type_left: break;
      sprintf(number, "%c", sht->freeze.n + 'A');
      *buf = strapp(*buf, number);
      *buf = strapp(*buf, "1");
      break;
  }
  *buf = strapp(*buf, "\" activePane=\"");
  switch (sht->freeze.type)
  {
    case libo_xl_freeze_type_none: break;
    case libo_xl_freeze_type_top:
      *buf = strapp(*buf, "bottomLeft");
      break;
    case libo_xl_freeze_type_left: break;
      *buf = strapp(*buf, "topRight");
      break;
  }
  *buf = strapp(*buf, "\" state=\"frozen\"/>");
  *buf = strapp(*buf, "<selection pane=\"");
  switch (sht->freeze.type)
  {
    case libo_xl_freeze_type_none: break;
    case libo_xl_freeze_type_top:
      *buf = strapp(*buf, "bottomLeft");
      break;
    case libo_xl_freeze_type_left: break;
      *buf = strapp(*buf, "topRight");
      break;
  }
  *buf = strapp(*buf, "\"/>");
  *buf = strapp(*buf, "</sheetView>\n");
  *buf = strapp(*buf, "</sheetViews>\n");
}

 /**
  * @fn static void libo_xl_sheet_formatpr_add(libo *l,
  *                                                int sheet,
  *                                                char **buf)
  *
  * @brief adds XL formatpr information to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_formatpr_add(libo *l, int sheet, char **buf)
{
  char number[25];

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  memset(number, 0, 25);

    /*
      <sheetFormatPr defaultRowHeight="15" customHeight="1" x14ac:dyDescent="0.3"/>
    */

  *buf = strapp(*buf, "<sheetFormatPr");
  if (l->xl->book->sheet[sheet]->default_row_height)
  {
    *buf = strapp(*buf, " defaultRowHeight=\"");
    sprintf(number, "%g", l->xl->book->sheet[sheet]->default_row_height);
    *buf = strapp(*buf, number);
    *buf = strapp(*buf, "\"");
    *buf = strapp(*buf, " customHeight=\"1\"");
  }
  *buf = strapp(*buf," x14ac:dyDescent=\"0.3\"/>\n");
}

 /**
  * @fn static void libo_xl_sheet_cols_add(libo *l, int sheet, char **buf)
  *
  * @brief adds XL worksheet column attributes to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_cols_add(libo *l, int sheet, char **buf)
{
  char number[25];
  int i;
  libo_xl_column *col;
  libo_xl_sheet *sht;

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (!l->xl) return;
  if (!l->xl->book) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  memset(number, 0, 25);

  sht = l->xl->book->sheet[sheet];

  if (!sht->column) sht->column = libo_xl_sheet_columns_create_defaults(sht);
  if (!sht->column) return;

    /*
      <cols>
        <col min="1" max="1" width="16.88671875" bestFit="1" customWidth="1"/>
      </cols>
    */

  *buf = strapp(*buf,"<cols>\n");
  for (i = 0; i < sht->n_cols; i++)
  {
    col = sht->column[i];

    *buf = strapp(*buf, "<col min=\"");
    sprintf(number, "%d", i+1);
    *buf = strapp(*buf, number);
    *buf = strapp(*buf, "\" max=\"");
    *buf = strapp(*buf, number);
    *buf = strapp(*buf, "\" width=\"");
    sprintf(number, "%f", col->width);
    *buf = strapp(*buf, number);
    *buf = strapp(*buf, "\" bestfit=\"");
    sprintf(number, "%d", col->autowidth ? 1 : 0);
    *buf = strapp(*buf, number);
    *buf = strapp(*buf, "\" customWidth=\"1\"/>\n");
  }
  *buf = strapp(*buf,"</cols>\n");
}

 /**
  * @fn static void libo_xl_sheet_sheetdata_add(libo *l,
  *                                                 int sheet,
  *                                                 char **buf)
  *
  * @brief adds XL worksheet data to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_sheetdata_add(libo *l, int sheet, char **buf)
{
  int i;

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

    /*
      <sheetData>
        ROWS
      </sheetData>
    */

  *buf = strapp(*buf, "<sheetData>\n");
  for (i = 0; i < l->xl->book->sheet[sheet]->n_rows; i++)
    libo_xl_sheet_sheetdata_row_add(l, sheet, i, buf);
  *buf = strapp(*buf, "</sheetData>\n");
}

 /**
  * @fn static void libo_xl_sheet_sheetdata_row_add(libo *l,
  *                                                     int sheet,
  *                                                     int row,
  *                                                     char **buf)
  *
  * @brief adds XL worksheet row data to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param row - index of row in sheet
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_sheetdata_row_add(libo *l, int sheet, int row, char **buf)
{
  int i;
  char number[25];

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (!l->xl->book) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  memset(number, 0, 25);

    /*
      NOTE:  Not sure what s="1" is, selected?
      <row r="1" spans="1:7" s="1" customFormat="1" ht="15" customHeight="1" x14ac:dyDescent="0.3">
        COLS
      </row>
    */

  *buf = strapp(*buf, "<row r=\"");
  sprintf(number, "%d", row+1);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, "\" spans=\"1:");
  sprintf(number, "%d", l->xl->book->sheet[sheet]->n_cols);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, "\" customFormat=\"1\" ht=\"");
  sprintf(number, "%g", l->xl->book->sheet[sheet]->default_row_height);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, "\" customHeight=\"1\" x14ac:dyDescent=\"0.3\">\n");
  for (i = 0; i < l->xl->book->sheet[sheet]->row[row]->n_cells; i++)
    libo_xl_sheet_sheetdata_row_col_add(l, sheet, row, i, buf);
  *buf = strapp(*buf, "</row>\n");
}

 /**
  * @fn static void libo_xl_sheet_sheetdata_row_col_add(libo *l,
  *                                                         int sheet,
  *                                                         int row,
  *                                                         int col,
  *                                                         char **buf)
  *
  * @brief adds XL worksheet cell data to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param row - index of row in sheet
  * @param col - index of col in row
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_sheetdata_row_col_add(libo *l,
                                                    int sheet,
                                                    int row,
                                                    int col,
                                                    char **buf)
{
  libo_xl_cell *cell;
  libo_xl_sheet *sht;
  char number[25];

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (!l->xl->book) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  sht = l->xl->book->sheet[sheet];
  if (!sht) return;

  if (row >= sht->n_rows) return;
  if (col >= sht->n_cols) return;

  cell = sht->row[row]->cell[col];

    /*
      <c r="A1" s="1" t="s"> //shared strings id
        <v>0</v>
      </c>
      <c r="C2" s="2"> //direct value (number)
        <v>156057</v>
      </c>
    */
  *buf = strapp(*buf, "<c r=\"");
  *buf = strapp(*buf, column_number_to_reference(col));
  sprintf(number, "%d", row+1);
  *buf = strapp(*buf, number);

  switch (cell->type)
  {
    case libo_xl_cell_type_none: break;

    case libo_xl_cell_type_reference:
      *buf = strapp(*buf, "\" s=\"1\" t=\"s\">\n");
      break;

#warning finish this
    case libo_xl_cell_type_expression:
      *buf = strapp(*buf, ">\n");
      break;

    case libo_xl_cell_type_number:
      *buf = strapp(*buf, "\" s=\"2\">\n");
      break;
  }

  *buf = strapp(*buf, "<v>");

  switch (cell->type)
  {
    case libo_xl_cell_type_none: break;

    case libo_xl_cell_type_reference:
      sprintf(number, "%d", cell->reference);
      *buf = strapp(*buf, number);
      //*buf = strapp(*buf, "\n");
      break;

#warning must handle expression type here
    case libo_xl_cell_type_expression: break;

    case libo_xl_cell_type_number:
      sprintf(number, "%g", cell->number);
      *buf = strapp(*buf, number);
      //*buf = strapp(*buf, "\n");
      break;
  }

  *buf = strapp(*buf, "</v>\n");
  *buf = strapp(*buf, "</c>\n");
}

 /**
  * @fn static void libo_xl_sheet_filter_add(libo *l, int sheet, char **buf)
  *
  * @brief adds XL worksheet filter information to XML buffer
  *
  * @param l - pointer to existing @a libo struct
  * @param sheet - index of sheet in book
  * @param buf - pointer to string holding XML buffer
  *
  * @par Returns
  * Nothing.
  */

static void libo_xl_sheet_filter_add(libo *l, int sheet, char **buf)
{
  libo_xl_sheet *sht;
  char number[25];

  if (!l) return;
  if (l->type != libo_type_xl) return;
  if (!l->xl->book) return;
  if (sheet >= l->xl->book->n_sheets) return;
  if (!buf) return;

  sht = l->xl->book->sheet[sheet];
  if (!sht) return;
  if (!sht->filter) return;

    /*
      <autoFilter ref="A1:E6" xr:uid="{00000000-0009-0000-0000-000000000000}">
        <sortState xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" ref="A2:E6">
          <sortCondition ref="A2:A6"/>
        </sortState>
      </autoFilter>
    */

  *buf = strapp(*buf, "<autoFilter ref=\"");
  *buf = strapp(*buf, column_number_to_reference(sht->filter->first_column));
  sprintf(number, "%d", 1);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, ":");
  *buf = strapp(*buf, column_number_to_reference(sht->filter->last_column));
  sprintf(number, "%d", sht->n_rows);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, "\" xr:uid=\"{00000000-0009-0000-0000-000000000000}\">");
  *buf = strapp(*buf, "<sortState xmlns:xlrd2=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" ref=\"");
  *buf = strapp(*buf, column_number_to_reference(sht->filter->first_column));
  sprintf(number, "%d", 2);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, ":");
  *buf = strapp(*buf, column_number_to_reference(sht->filter->last_column));
  sprintf(number, "%d", sht->n_rows);
  *buf = strapp(*buf, number);
  *buf = strapp(*buf, "\">");
  *buf = strapp(*buf, "</sortState>");
  *buf = strapp(*buf, "</autoFilter>");
}

  /**
   *  @fn void libo_xl_strings_count_action(avl_node *n)
   *
   *  @brief called by avl_walk() when counting XL strings
   *
   *  @param n - pointer to existing @a avl_node
   *
   *  @par Returns
   *  Nothing.
   */

static void libo_xl_strings_count_action(avl_node *n) { ++_strings_count; }

  /**
   *  @fn void libo_xl_strings_add_action(avl_node *n)
   *
   *  @brief called by avl_walk() when adding XML content to strings
   *
   *  @param n - pointer to existing @a avl_node
   *
   *  @par Returns
   *  Nothing.
   */

static void libo_xl_strings_add_action(avl_node *n)
{
  string *str;

  if (!n) return;

  str = (string *)n->value;
  if (!str) return;

  _strings_buf = strapp(_strings_buf, "<si>");
  _strings_buf = strapp(_strings_buf, "<t>");
  _strings_buf = strapp(_strings_buf, str->text);
  _strings_buf = strapp(_strings_buf, "</t>");
  _strings_buf = strapp(_strings_buf, "</si>");
}

  /**
   *  @fn char *column_number_to_reference(unsigned int n)
   *
   *  @brief converts a column number to XL cell reference
   *
   *  @param n - column number
   *
   *  @return string containing XL cell reference
   */

static char *column_number_to_reference(unsigned int n)
{
  static char ref[10];
  int i = 0;
  unsigned int r;

  memset(ref, 0, 10);

  ++n;

  do
  {
    r = (n - 1) % 26;
    n = (n - 1) / 26;
    ref[i] = 'A' + r;
    ++i;
  } while (n);

  return reverse(ref);
}

  /**
   *  @fn char *reverse(char *s)
   *
   *  @brief reverses order of characters in a string
   *
   *  NOTE:  It is the caller's responsibility to free allocated memory
   *         of the reversed string
   *
   *  @param s - string to reverse
   *
   *  @return reversed string
   */

static char *reverse(char *s)
{
  int len;
  int i, j;
  char *t;

  if (!s) return NULL;

  len = strlen(s);
  if (len < 2) return s;

  t = strdup(s);
  memset(t, 0, len);

  i = 0;
  j = len - 1;
  while (i < len) t[i++] = s[j--];

  strcpy(s, t);
  free(t);

  return s;
}

  /**
   *  @fn void libo_xl_row_fill(libo_xl_sheet *sheet, int max_row)
   *
   *  @brief fills XL worksheet with rows up to @p max_row
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *  @param max_row - index of last row to create
   *
   *  @par Returns
   *  Nothing.
   */

static void libo_xl_row_fill(libo_xl_sheet *sheet, int max_row)
{
  int i;

  if (!sheet) return;

  if (sheet->row)
    sheet->row = (libo_xl_row **)realloc(sheet->row, sizeof(libo_xl_row *) * max_row);
  else
    sheet->row = (libo_xl_row **)malloc(sizeof(libo_xl_row *) * max_row);

  for (i = sheet->n_rows; i < max_row; i++)
    sheet->row[i] = libo_xl_row_new();

  sheet->n_rows = max_row;

  return;
}

  /**
   *  @fn void libo_xl_col_fill(libo_xl_sheet *sheet,
   *                            int row,
   *                            int max_col)
   *
   *  @brief fills XL row of worksheet with cells up to @p max_col
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *  @param row - index of row to fill
   *  @param max_col - index of last col to create
   *
   *  @par Returns
   *  Nothing.
   */

static void libo_xl_col_fill(libo_xl_sheet *sheet, int row, int max_col)
{
  int i;

  if (!sheet) return;

  if (row >= sheet->n_rows) libo_xl_row_fill(sheet, row);

  if (sheet->row[row]->cell)
    sheet->row[row]->cell = (libo_xl_cell **)realloc(sheet->row[row]->cell, sizeof(libo_xl_cell *) * max_col);
  else
    sheet->row[row]->cell = (libo_xl_cell **)malloc(sizeof(libo_xl_cell *) * max_col);

  for (i = sheet->row[row]->n_cells; i < max_col; i++)
    sheet->row[row]->cell[i] = libo_xl_cell_new();

  sheet->row[row]->n_cells = max_col;

  return;
}

  /**
   *  @fn void string_dumper(avl_node *n)
   *
   *  @brief outputs @a string struct contents for "dumpers"
   *
   *  @param n - pointer to existing @a avl_node struct
   *
   *  @par Returns
   *  Nothing.
   */

static void string_dumper(avl_node *n)
{
  string *str = NULL;

  if (!n) return;

  str = (string *)n->value;

  do_indent(_dumper_file, _strings_indent);
  fprintf(_dumper_file,
          "id=%d, text='%s'\n",
          str->id,
          str->text ? str->text : "");
}

  /**
   *  @fn static void libo_xl_cell_clear(libo_xl_cell *cell)
   *
   *  @brief clears contents of @p cell, freeing memory if needed
   *
   *  @param cell - pointer to existing @a libo_xl_cell struct
   *
   *  @par Returns
   *  Nothing.
   */

static void libo_xl_cell_clear(libo_xl_cell *cell)
{
  if (!cell) return;

  switch (cell->type)
  {
    case libo_xl_cell_type_reference:
      cell->reference = 0;
      break;

    case libo_xl_cell_type_expression:
      if (cell->expression.value) free(cell->expression.value);
      if (cell->expression.formula) free(cell->expression.formula);
      memset(&cell->expression, 0, sizeof(libo_xl_cell_expression));
      break;

    case libo_xl_cell_type_number:
      cell->reference = 0;
      break;

    case libo_xl_cell_type_none:
    default:
      break;
  }

  libo_xl_cell_set_type(cell, libo_xl_cell_type_none);
}

  /**
   *  @fn void libo_xl_sheet_count_columns(libo_xl_sheet *xls)
   *
   *  @brief finds maximum columns in @p xls
   *
   *  NOTE:  This function modifies the n_cols field in @p xls
   *
   *  @param xls - pointer to existing @a libo_xl_sheet struct
   *
   *  @par Returns
   *  Nothing.
   */

static void libo_xl_sheet_count_columns(libo_xl_sheet *xls)
{
  int i;
  libo_xl_row *row;

  if (!xls) return;

  xls->n_cols = 0;

  for (i = 0; i < xls->n_rows; i++)
  {
    row = xls->row[i];
    if (row->n_cells > xls->n_cols)
      xls->n_cols = row->n_cells;
  }
}

  /**
   *  @fn libo_xl_column **libo_xl_sheet_columns_create_defaults(libo_xl_sheet *sheet)
   *
   *  @brief creates a @a libo_xl_column for each column used in @p sheet
   *
   *  @param sheet - pointer to existing @a libo_xl_sheet struct
   *
   *  @return pointer to array of @a libo_xl_column pointers
   */

static libo_xl_column **libo_xl_sheet_columns_create_defaults(libo_xl_sheet *sheet)
{
  libo_xl_column **columns = NULL;
  int i;

  if (!sheet) goto exit;

  columns = realloc(columns, sizeof(libo_xl_column *) * sheet->n_cols);
  if (!columns) goto exit;

  for (i = 0; i < sheet->n_cols; i++)
    columns[i] = libo_xl_column_new_with_values(15, 1);

exit:
  return columns;
}

