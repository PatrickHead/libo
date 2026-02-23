AM_CPPFLAGS = -I $(srcdir)/include

AM_CFLAGS = -O3 -g0 -Wall

ARFLAGS = cr

lib_LIBRARIES = lib/libo.a
lib_libo_a_SOURCES = src/libo.c include/libo.h
lib_libo_a_CFLAGS = -O3 -g0 -Wall @LIBXML2_CFLAGS@ @LIBSTRINGS_CFLAGS@ @LIBCSV_CFLAGS@ @LIBZIP_CFLAGS@

bin_PROGRAMS = bin/test-libo
bin_test_libo_SOURCES = src/test-libo.c
bin_test_libo_CFLAGS = -O3 -g0 -Wall @LIBXML2_CFLAGS@ @LIBSTRINGS_CFLAGS@ @LIBCSV_CFLAGS@ @LIBZIP_CFLAGS@
bin_test_libo_LDADD = @LIBCSV_LIBS@ lib/libo.a @LIBXML2_LIBS@ @LIBSTRINGS_LIBS@ @LIBZIP_LIBS@ @AVL_LIBS@

include_HEADERS = include/libo.h

EXTRA_DIST = windows extras acdoxygen.m4 amdoxygen.am doxygen.o.cfg o.pc .gitignore

super-clean: clean distclean
	@find . -name 'Makefile.in' -delete
	@find . -name 'config.h.in' -delete
	@find . -name '*.tar' -delete
	@find . -name '*.tar.*' -delete
	@find . -name '*~' -delete
	@find . -name '*.obj' -delete
	@find . -name '*.lib' -delete
	@find . -name '*.exe' -delete
	@rm -rf autom4te.cache
	@rm -f aclocal.m4
	@rm -f configure
	@rm -rf src/.deps
	@rm -f texput.log

include amdoxygen.am

MOSTLYCLEANFILES = $(DX_CLEANFILES)

EXTRA_DIST += $(DX_CONFIG)

install-data-hook:
	@if test -d $(DX_HTML_DIR); \
	 then \
	   echo installing $(DX_HTML_DIR) to $(prefix)/share/doc/libo; \
	   mkdir -p $(prefix)/share/doc/libo; \
	   cp -r $(DX_HTML_DIR) $(prefix)/share/doc/libo; \
	 fi
	@echo installing libo.pc
	@mkdir -p $(prefix)/share/pkgconfig
	@cp libo.pc $(prefix)/share/pkgconfig/libo.pc

uninstall-hook:
	@rm -rf $(prefix)/share/doc/libo
	@rm -rf $(prefix)/share/pkgconfig/libo.pc

clean-local:
	@(cd windows; make clean)

code-stats:
	@echo -n " total lines of source code: "; wc -l src/*.[ch] include/*.[ch] | tail -1 | sed -e 's/ total//' -e 's/^  *//'
	@echo -n "actual lines of source code: "; (gcc -fpreprocessed -dD -E src/*.[ch] include/*.[ch] | sed -e '/^$$/d' -e '/^# [0-9]/d' | wc -l)

