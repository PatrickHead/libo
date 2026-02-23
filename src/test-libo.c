#include <stdlib.h>
#include <stdio.h>
#include <getopt.h>

#include "libo.h"

typedef enum
{
  DUMP,
  API
} test_type;

static libo *test_creation_functions(void);

int main(int argc, char **argv)
{
  libo *l;
  libo_xl *xl;
  libo_xl_book *book;
  libo_xl_sheet *sheet;
  libo_xl_row *row;
  libo_xl_cell *cell;
  int c;
  test_type mode = API;
  int sheet_count;
  char *sheet_name;
  int sheet_id;
  char *sheet_rid;
  int row_count;
  int column_count;
  int cell_count;
  libo_xl_cell_type cell_type;
  int cell_reference;
  char *cell_text;
  libo_xl_cell_expression *cell_expression;
  char *cell_formula;
  char *cell_value;
  double cell_number;
  int i, j, k;
  char *sv;

  while ((c = getopt(argc, argv, "da")) != EOF)
  {
    switch (c)
    {
      case 'd':
        mode = DUMP;
        break;
      case 'a':
        mode = API;
        break;
      default:
        return 1;
        break;
    }
  }

  libo_init();

  printf("\n\nStarting READ and DUMP Tests\n\n");

  l = libo_open("xlsx/all.xlsx");
  if (!l)
    return 1;

  if (mode == DUMP)
    libo_dump(l, stdout, 0);
  else
  {
    printf("API TESTS:\n");
    printf("libo_get_type(%p)=%d\n", l, libo_get_type(l));
    printf("libo_get_path(%p)=%s\n", l, libo_get_path(l));

    printf("libo_get_xl(%p)=%p\n", l, xl = libo_get_xl(l));
    printf("libo_xl_get_book(%p)=%p\n", xl, book = libo_xl_get_book(xl));
    printf("libo_xl_book_get_sheet_count(%p)=%d\n", book, sheet_count = libo_xl_book_get_sheet_count(book));

    for (i = 0; i < sheet_count; i++)
    {
      printf("libo_xl_book_get_sheet(%p, %d)=%p\n", book, i, sheet = libo_xl_book_get_sheet(book, i));
      printf("libo_xl_sheet_get_row_count(%p)=%d\n", sheet, row_count = libo_xl_sheet_get_row_count(sheet));
      printf("libo_xl_sheet_get_column_count(%p)=%d\n", sheet, column_count = libo_xl_sheet_get_column_count(sheet));
      printf("libo_xl_sheet_get_name(%p)=%s\n", sheet, sheet_name = libo_xl_sheet_get_name(sheet));
      printf("libo_xl_sheet_get_id(%p)=%d\n", sheet, sheet_id = libo_xl_sheet_get_id(sheet));
      printf("libo_xl_sheet_get_rid(%p)=%s\n", sheet, sheet_rid = libo_xl_sheet_get_rid(sheet));

      for (j = 0; j < row_count; j++)
      {
        printf("libo_xl_sheet_get_row(%p, %d)=%p\n", sheet, j, row = libo_xl_sheet_get_row(sheet, j));
        printf("libo_xl_row_get_cell_count(%p)=%d\n", row, cell_count = libo_xl_row_get_cell_count(row));

        for (k = 0; k < cell_count; k++)
        {
          printf("libo_xl_row_get_cell(%p, %d)=%p\n", row, k, cell = libo_xl_row_get_cell(row, k));
          printf("libo_xl_cell_get_type(%p)=%d\n", cell, cell_type = libo_xl_cell_get_type(cell));

          switch (cell_type)
          {
            case libo_xl_cell_type_none:
              printf("UNKNOWN CELL TYPE\n");
              break;
            case libo_xl_cell_type_reference:
              printf("libo_xl_cell_get_reference(%p)=%d\n", cell, cell_reference = libo_xl_cell_get_reference(cell));
              printf("libo_xl_cell_get_text(%p, %p)=%s\n",
                     xl,
                     cell,
                     (cell_text = libo_xl_cell_get_text(xl, cell)) ? cell_text : "[NONE]");
              fflush(stdout);
              break;
            case libo_xl_cell_type_expression:
              printf("libo_xl_cell_get_expression(%p)=%p\n", cell, cell_expression = libo_xl_cell_get_expression(cell));
              printf("libo_xl_cell_expression_get_formula(%p)=%s\n", cell_expression, cell_formula = libo_xl_cell_expression_get_formula(cell_expression));
              printf("libo_xl_cell_expression_get_value(%p)=%s\n", cell_expression, cell_value = libo_xl_cell_expression_get_value(cell_expression));
              break;
            case libo_xl_cell_type_number:
              printf("libo_xl_cell_get_number(%p)=%f\n", cell, cell_number = libo_xl_cell_get_number(cell));
              break;
          }

          printf("libo_xl_cell_get_string_value(%p, %p)=%s\n",
                 xl,
                 cell,
                 (sv = libo_xl_cell_get_string_value(xl, cell)) ? sv : "[NONE]");
          fflush(stdout);
          if (sv) free(sv);
        }
      }
    }
  }

  libo_close(l);

  printf("\n\nREAD and DUMP Tests Complete\n\n");

  printf("\n\nStarting CREATION Tests\n\n");

  l = test_creation_functions();
  printf("l=%p\n", l);

  libo_dump(l, stdout, 0);

  libo_write(l, l->path);

  printf("\n\nCREATION Tests Complete\n\n");

  libo_cleanup();

  return 0;
}

static libo *test_creation_functions(void)
{
  libo *doc = NULL;
  libo_xl *xl = NULL;
  libo_xl_book *book = NULL;
  libo_xl_sheet *sheet = NULL;
  libo_xl_row *row = NULL;
  libo_xl_cell *cell = NULL;
  char name[25];
  int i;
  int j;

  doc = libo_new();
  if (!doc) goto exit;

  libo_set_type(doc, libo_type_xl);
  libo_set_path(doc, "TEST-CREATION.xlsx");

  remove(doc->path);

  xl = libo_get_xl(doc);
  if (!xl) goto exit;

    /* The following shouldn't be needed, but libo doesn't do this in a convenient way */
  book = xl->book = libo_xl_book_new();

  for (i = 0; i < 10; i++)
  {
      /* create new sheet */

    sheet = libo_xl_sheet_new();
    sprintf(name, "Sheet%d", libo_xl_book_get_sheet_count(book) + 1);
    libo_xl_sheet_set_name(sheet, name);

        /* Create header row */

    row = libo_xl_row_new();
    if (!row) goto loopend;

    cell = libo_xl_cell_new();
    if (!cell) goto loopend;

    libo_xl_cell_set_text(xl, cell, "Datum");
    
    libo_xl_row_add(row, cell);
    libo_xl_cell_free(cell); cell = NULL;

    libo_xl_sheet_add(sheet, row);
    libo_xl_row_free(row); row = NULL;

        /* Create subsequent data rows */
    for (j = 0; j < 12; j++)
    {
      row = libo_xl_row_new();
      if (!row) goto loopend;

      cell = libo_xl_cell_new();
      if (!cell) goto loopend;

      libo_xl_cell_set_number(cell, (double)((i+1) * 10 + j));

      libo_xl_row_add(row, cell);
      libo_xl_cell_free(cell); cell = NULL;

      libo_xl_sheet_add(sheet, row);
      libo_xl_row_free(row); row = NULL;
    }

    libo_xl_book_add(book, sheet);
    libo_xl_sheet_free(sheet); sheet = NULL;

loopend:
    if (row) libo_xl_row_free(row);
    if (cell) libo_xl_cell_free(cell);
    if (sheet) libo_xl_sheet_free(sheet);
  }

exit:
  return doc;
}

