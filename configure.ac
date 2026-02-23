#                                               -*- Autoconf -*-
# Process this file with autoconf to produce a configure script.

AC_PREREQ([2.71])
AC_INIT([libo], [1.0.0], [patrickhead@gmail.com])
AC_CONFIG_MACRO_DIR([.])
AC_CONFIG_AUX_DIR(tools)
AC_CONFIG_SRCDIR([src/libo.c])
AC_CONFIG_HEADERS([src/config.h])

# Doxygen features and initialization

DX_HTML_FEATURE(ON)
DX_CHM_FEATURE(OFF)
DX_CHI_FEATURE(OFF)
DX_MAN_FEATURE(OFF)
DX_RTF_FEATURE(OFF)
DX_XML_FEATURE(OFF)
DX_PDF_FEATURE(ON)
DX_PS_FEATURE(OFF)
DX_INIT_DOXYGEN(libo, doxygen.libo.cfg)

AM_INIT_AUTOMAKE([subdir-objects])
AM_SILENT_RULES([yes])

# Checks for programs.
AC_PROG_CC
AC_PROG_RANLIB

# Require pkg-config
PKG_PROG_PKG_CONFIG

# Check for libxml2 library
PKG_CHECK_MODULES([LIBXML2], [libxml-2.0 >= 2.9.0],
  [AC_DEFINE([HAVE_LIBXML2], [1], [Define if libxml2 is available])],
  [AC_MSG_ERROR([libxml2 not found. Install libxml2 library.])]
)

# Check for libzip library
PKG_CHECK_MODULES([LIBZIP], [libzip >= 1.7.0],
  [AC_DEFINE([HAVE_LIBZIP], [1], [Define if libzip is available])],
  [AC_MSG_ERROR([libzip not found. Install libzip library.])]
)

# Check for avl library
PKG_CHECK_MODULES([AVL], [avl >= 1.0.0],
  [AC_DEFINE([HAVE_AVL], [1], [Define if avl is available])],
  [AC_MSG_ERROR([avl not found. Install avl library.])]
)

# Check for strings library
PKG_CHECK_MODULES([LIBSTRINGS], [libstrings >= 1.0.0],
  [AC_DEFINE([HAVE_LIBSTRINGS], [1], [Define if libstrings is available])],
  [AC_MSG_ERROR([libstrings not found. Install libstrings library.])]
)

# Check for csv library
PKG_CHECK_MODULES([LIBCSV], [libcsv >= 1.0.0],
  [AC_DEFINE([HAVE_LIBCSV], [1], [Define if libcsv is available])],
  [AC_MSG_ERROR([libcsv not found. Install libcsv library.])]
)

# Checks for header files.
AC_CHECK_HEADERS([unistd.h])

# Checks for typedefs, structures, and compiler characteristics.
AC_TYPE_SIZE_T
AC_TYPE_UINT8_T

# Checks for library functions.
AC_FUNC_MALLOC
AC_FUNC_REALLOC
AC_CHECK_FUNCS([getcwd memset mkdir strcasecmp strdup strncasecmp strrchr])

AC_CONFIG_FILES([Makefile libo.pc])

AC_OUTPUT
