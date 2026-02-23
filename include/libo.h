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
 *  @file libo.h
 *  @brief  Header file for libo
 *
 *  A library to aid in manipulating data in Office files.
 *
 *  NOTE:  Currently only XLSX (Excel) files are supported
 */

#ifndef LIBO_H
#define LIBO_H

#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxml/xpath.h>
#include <libxml/xpathInternals.h>
#include <zip.h>

#include <libstrings.h>

#define VERSION "1.0.0"  /**<  Version of libo  */

  /**
   *  @typedef enum libo_type
   *
   *  @brief Office document types
   */

typedef enum
{
  libo_type_none,  /**<  no or unknown document type  */
  libo_type_xl,    /**<  Excel                        */
  libo_type_doc,   /**<  Word                         */
  libo_type_pp     /**<  PowerPoint                   */
} libo_type;

  /**
   *  @typedef enum libo_xl_cell_type
   *
   *  @brief type of data stored in an Excel cell
   */

typedef enum
{
  libo_xl_cell_type_none,        /**<  no or unknown cell type  */
  libo_xl_cell_type_reference,   /**<  reference  */
  libo_xl_cell_type_expression,  /**<  expression, such as formula  */
  libo_xl_cell_type_number       /**<  direct value  */
} libo_xl_cell_type;

  /**
   *  @typedef enum libo_xl_expression_type
   *
   *  @brief type of data stored in an Excel expression
   */

typedef enum
{
  libo_xl_expression_type_none,
  libo_xl_expression_type_formula,
  libo_xl_expression_type_value
} libo_xl_expression_type;

  /**
   *  @typedef struct libo_xl_cell_expression libo_xl_cell_expression
   *
   *  @brief create a type for struct @a libo_xl_cell_expression
   */

typedef struct libo_xl_cell_expression libo_xl_cell_expression;

  /**
   *  @struct libo_xl_cell_expression
   *
   *  @brief struct that holds an expression
   */

struct libo_xl_cell_expression
{
  char *formula;  /**<  formula used to calculate cell's value  */
  char *value;    /**<  value produced by formula               */
};

  /**
   *  @typedef struct libo_xl_cell libo_xl_cell
   *
   *  @brief create a type for struct @a libo_xl_cell
   */

typedef struct libo_xl_cell libo_xl_cell;

  /**
   *  @struct libo_xl_cell
   *
   *  @brief struct that holds Excel cell data
   */

struct libo_xl_cell
{
  libo_xl_cell_type type;      /**<  type of cell, see @a libo_xl_cell_type  */
  union
  {
    int reference;                       /**<  reference identifier          */
    libo_xl_cell_expression expression;  /**<  expression                    */
    double number;                       /**<  direct value                  */
  };
};

  /**
   *  @typedef struct libo_xl_row libo_xl_row;
   *
   *  @brief create a type for struct @a libo_xl_row
   */

typedef struct libo_xl_row libo_xl_row;

  /**
   *  @struct libo_xl_row
   *
   *  @brief struct that holds an Excel row of cells
   */

struct libo_xl_row
{
  int n_cells;          /**<  number of cells in row  */
  libo_xl_cell **cell;  /**<  array of cells          */
};

  /**
   *  @typedef enum libo_xl_freeze_type
   *
   *  @brief type of freeze for xl rows/columns
   */

typedef enum
{
  libo_xl_freeze_type_none,
  libo_xl_freeze_type_top,
  libo_xl_freeze_type_left
} libo_xl_freeze_type;

  /**
   *  @typedef struct libo_xl_freeze libo_xl_freeze;
   *
   *  @brief create a type for struct @a libo_xl_freeze
   */

typedef struct libo_xl_freeze libo_xl_freeze;

  /**
   *  @struct libo_xl_freeze
   *
   *  @brief struct that holds definition for Excel row/column freeze
   */

struct libo_xl_freeze
{
  libo_xl_freeze_type type;  /**<  @a libo_xl_freeze_type of freeze       */
  int n;                     /**<  number of rows or columents to freeze  */
};

  /**
   *  @typedef struct libo_xl_column libo_xl_column;
   *
   *  @brief create a type for struct @a libo_xl_column
   */

typedef struct libo_xl_column libo_xl_column;

  /**
   *  @struct libo_xl_column
   *
   *  @brief struct that holds attributes for Excel column
   */

struct libo_xl_column
{
  float width;    /**<  absolute width of column  */
  int autowidth;  /**<  auto calculate width of column based on contents  */
};

  /**
   *  @typedef struct libo_xl_filter libo_xl_filter;
   *
   *  @brief create a type for struct @a libo_xl_filter
   */

typedef struct libo_xl_filter libo_xl_filter;

  /**
   *  @struct libo_xl_filter
   *
   *  @brief struct that holds filter attributes
   */

struct libo_xl_filter
{
  unsigned int first_column;  /**<  first column to filter  */
  unsigned int last_column;   /**<  last column to filter   */
};

  /**
   *  @typedef struct libo_xl_sheet libo_xl_sheet;
   *
   *  @brief create a type for struct @a libo_xl_sheet
   */

typedef struct libo_xl_sheet libo_xl_sheet;

  /**
   *  @struct libo_xl_sheet
   *
   *  @brief struct that holds an Excel work sheet
   */

struct libo_xl_sheet
{
  int n_rows;                 /**<  number of rows in sheet             */
  int n_cols;                 /**<  maximum number of columns in sheet  */
  char *name;                 /**<  title of sheet                      */
  int ID;                     /**<  identifier of sheet                 */
  char *rID;                  /**<  reference identifier                */
  double default_row_height;  /**<  default height of each row          */
  libo_xl_freeze freeze;      /**<  row/column freeze setting           */
  libo_xl_row **row;          /**<  arrow of rows                       */
  libo_xl_column **column;    /**<  columnn attributes                  */
  libo_xl_filter *filter;     /**<  filtered columns                    */
};

  /**
   *  @typedef struct libo_xl_book libo_xl_book;
   *
   *  @brief create a type for struct @a libo_xl_book
   */

typedef struct libo_xl_book libo_xl_book;

  /**
   *  @struct libo_xl_book
   *
   *  @brief struct that holds an Excel workbook
   */

struct libo_xl_book
{
  int n_sheets;           /**<  number of worksheets  */
  libo_xl_sheet **sheet;  /**<  array of work sheets  */
};

  /**
   *  @typedef struct libo_xl libo_xl;
   *
   *  @brief create a type for struct @a libo_xl
   */

typedef struct libo_xl libo_xl;

  /**
   *  @struct libo_xl
   *
   *  @brief struct that holds an Excel document
   */

struct libo_xl
{
  libo_xl_book *book;  /**< workbook           */
  strings *strings;    /**< strings dictionary */
};

  /**
   *  @typedef typedef struct libo_doc libo_doc;
   *
   *  @brief create a type for struct @a libo_doc
   */

typedef struct libo_doc libo_doc;

  /**
   *  @struct libo_doc
   *
   *  @brief struct that holds a Word document
   *
   *  NOTE:  libo_doc is NOT implemented
   */

struct libo_doc
{
  // NOT IMPLEMENTED
};

  /**
   *  @typedef struct libo_pp libo_pp;
   *
   *  @brief create a type for struct @a libo_pp
   */

typedef struct libo_pp libo_pp;

  /**
   *  @struct libo_pp
   *
   *  @brief struct that holds a PowerPoint document
   *
   *  NOTE:  libo_pp is NOT implemented
   */

struct libo_pp
{
  // NOT IMPLEMENTED
};

  /**
   *  @typedef struct libo libo;
   *
   *  @brief create a type for struct @a libo
   */

typedef struct libo libo;

  /**
   *  @struct libo
   *
   *  @brief struct that holds an Office document
   */

struct libo
{
  char *path;      /**<  full path to document file       */
  libo_type type;  /**<  type of Office document          */
  zip_t *z;        /**<  ZIP file data                    */
  union
  {
    libo_xl *xl;    /**<  pointer to Excel document       */
    libo_doc *doc;  /**<  pointer to Word document       .*/
    libo_pp *pp;    /**<  pointer to PowerPoint document  */
  };
};

  /*
   *  Library helpers
   */

void libo_init(void);
void libo_cleanup(void);

  /*
   *  LIBO
   */

libo *libo_new(void);
libo *libo_dup(libo *l);
libo *libo_open(char *path);
void libo_free(libo *l);
void libo_close(libo *l);

libo_type libo_get_type(libo *l);
void libo_set_type(libo *l, libo_type type);

char *libo_get_path(libo *l);
void libo_set_path(libo *l, char *path);

libo_xl *libo_get_xl(libo *l);
libo_doc *libo_get_doc(libo *l);
libo_pp *libo_get_pp(libo *l);

void libo_dump(libo *l, FILE *stream, int indent);

int libo_write(libo *l, char *path);

  /*
   *  DOC
   */

libo_doc *libo_doc_new(void);
libo_doc *libo_doc_dup(libo_doc *doc);
void libo_doc_free(libo_doc *doc);
void libo_doc_dump(libo_doc *doc, FILE *stream, int indent);

  /*
   *  PP
   */

libo_pp *libo_pp_new(void);
libo_pp *libo_pp_dup(libo_pp *pp);
void libo_pp_free(libo_pp *pp);
void libo_pp_dump(libo_pp *pp, FILE *stream, int indent);

  /*
   *  XL
   */

libo_xl *libo_xl_new(void);
libo_xl *libo_xl_dup(libo_xl *xl);
libo_xl *libo_xl_read(libo *l);
void libo_xl_free(libo_xl *xl);

libo_xl_book *libo_xl_get_book(libo_xl *xl);

strings *libo_xl_strings_read(libo *l);
void libo_xl_strings_dump(strings *strs, FILE *stream, int indent);

void libo_xl_dump(libo_xl *xl, FILE *stream, int indent);

  /*
   *  XL book
   */

libo_xl_book *libo_xl_book_new(void);
libo_xl_book *libo_xl_book_dup(libo_xl_book *book);
libo_xl_book *libo_xl_book_read(libo *l);
void libo_xl_book_free(libo_xl_book *book);

int libo_xl_book_get_sheet_count(libo_xl_book *xlb);
libo_xl_sheet *libo_xl_book_get_sheet(libo_xl_book *xlb, int n);

void libo_xl_book_add(libo_xl_book *xlb, libo_xl_sheet *xls);

void libo_xl_book_dump(libo_xl_book *lxb, FILE *stream, int indent);

  /*
   *  XL sheet
   */

libo_xl_sheet *libo_xl_sheet_new(void);
libo_xl_sheet *libo_xl_sheet_dup(libo_xl_sheet *sheet);
void libo_xl_sheet_read(libo *l, libo_xl_sheet *sheet, int n);
libo_xl_sheet *libo_xl_sheet_meta_read(xmlDocPtr doc, int n);
libo_xl_row **libo_xl_sheet_rows_read(libo_xl_sheet *sheet, xmlDocPtr doc);
void libo_xl_sheet_free(libo_xl_sheet *sheet);

int libo_xl_sheet_get_row_count(libo_xl_sheet *xls);
int libo_xl_sheet_get_column_count(libo_xl_sheet *xls);

libo_xl_row *libo_xl_sheet_get_row(libo_xl_sheet *xls, int n);

char *libo_xl_sheet_get_name(libo_xl_sheet *xls);
void libo_xl_sheet_set_name(libo_xl_sheet *xls, char *name);

int libo_xl_sheet_get_id(libo_xl_sheet *xls);
void libo_xl_sheet_set_id(libo_xl_sheet *xls, int id);

char *libo_xl_sheet_get_rid(libo_xl_sheet *xls);
void libo_xl_sheet_set_rid(libo_xl_sheet *xls, char *rid);

void libo_xl_sheet_set_default_row_height(libo_xl_sheet *sheet,
                                          double default_row_height);

libo_xl_freeze *libo_xl_sheet_get_freeze(libo_xl_sheet *sheet);
void libo_xl_sheet_set_freeze(libo_xl_sheet *sheet,
                              libo_xl_freeze_type type,
                              int n);

void libo_xl_sheet_add(libo_xl_sheet *xls, libo_xl_row *xlr);

void libo_xl_sheet_add_filter(libo_xl_sheet *sheet,
                              unsigned int first_column,
                              unsigned int last_column);
void libo_xl_sheet_remove_filter(libo_xl_sheet *sheet);

void libo_xl_sheet_dump(libo_xl_sheet *lxs, FILE *stream, int indent);

  /*
   *  XL column
   */

libo_xl_column *libo_xl_column_new(void);
libo_xl_column *libo_xl_column_new_with_values(float width, int autowidth);
void libo_xl_column_free(libo_xl_column *column);

  /*
   *  XL filter
   */

libo_xl_filter *libo_xl_filter_new(void);
libo_xl_filter *libo_xl_filter_new_with_values(unsigned int first_column,
                                               unsigned int last_column);
void libo_xl_filter_free(libo_xl_filter *filter);

  /*
   *  XL row
   */

libo_xl_row *libo_xl_row_new(void);
libo_xl_row *libo_xl_row_dup(libo_xl_row *row);
void libo_xl_row_free(libo_xl_row *row);

int libo_xl_row_get_cell_count(libo_xl_row *xlr);

libo_xl_cell *libo_xl_row_get_cell(libo_xl_row *xlr, int n);

void libo_xl_row_add(libo_xl_row *xlr, libo_xl_cell *xlc);

void libo_xl_row_dump(libo_xl_row *row, FILE *stream, int indent);

  /*
   *  XL cell
   */

libo_xl_cell *libo_xl_cell_new(void);
libo_xl_cell *libo_xl_cell_dup(libo_xl_cell *cell);
libo_xl_cell *libo_xl_cell_create(libo_xl_sheet *sheet, int row, int col);
void libo_xl_cell_free(libo_xl_cell *cell);

libo_xl_cell_type libo_xl_cell_get_type(libo_xl_cell *xlc);
void libo_xl_cell_set_type(libo_xl_cell *xlc, libo_xl_cell_type type);

char *libo_xl_cell_get_string_value(libo_xl *xl, libo_xl_cell *xlc);

int libo_xl_cell_get_reference(libo_xl_cell *xlc);
void libo_xl_cell_set_reference(libo_xl_cell *xlc, int reference);

char *libo_xl_cell_get_text(libo_xl *xl, libo_xl_cell *xlc);
void libo_xl_cell_set_text(libo_xl *xl, libo_xl_cell *xlc, char *text);

libo_xl_cell_expression *libo_xl_cell_get_expression(libo_xl_cell *xlc);
void libo_xl_cell_set_expression(libo_xl_cell *xlc, libo_xl_cell_expression *xlce);

char *libo_xl_cell_expression_get_formula(libo_xl_cell_expression *xlce);
void libo_xl_cell_expression_set_formula(libo_xl_cell_expression *xlce, char *formula);

char *libo_xl_cell_expression_get_value(libo_xl_cell_expression *xlce);
void libo_xl_cell_expression_set_value(libo_xl_cell_expression *xlce, char *value);

double libo_xl_cell_get_number(libo_xl_cell *xlc);
void libo_xl_cell_set_number(libo_xl_cell *xlc, double number);

void libo_xl_cell_dump(libo_xl_cell *cell, FILE *stream, int indent);

  /*
   *  Type helpers
   */

char *libo_xl_cell_type_to_string(libo_xl_cell_type ct);
char *libo_type_to_string(libo_type lt);

#endif //LIBO_H
