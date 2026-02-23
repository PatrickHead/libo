CC = x86_64-w64-mingw32-gcc
SRCDIR = ../src
INCLDIR = ../include

WINPKGSDIR = $(HOME)/develop/win-pkgs
WINPKGINCLDIR = $(WINPKGSDIR)/include
WINPKGLIBDIR = $(WINPKGSDIR)/lib
LIBXML2INCLDIR = $(HOME)/develop/gtk4-for-win/include/libxml2
LIBXML2LIBDIR = $(HOME)/develop/gtk4-for-win/lib

CC = x86_64-w64-mingw32-gcc
COPTS = -O3 -g0 -Wall -I $(INCLDIR) -I $(WINPKGINCLDIR) -I $(LIBXML2INCLDIR) -L . -L $(WINPKGLIBDIR) -L $(LIBXML2LIBDIR)
AR = x86_64-w64-mingw32-ar
RANLIB = x86_64-w64-mingw32-ranlib

all: o.lib test-libo.exe

test-libo.exe: test-libo.obj libo.a
	$(CC) $(COPTS) -L. -o test-libo.exe test-libo.obj o.lib -lcsv -lzip -lxml2 -lstrings -lavl

test-libo.obj: $(SRCDIR)/test-libo.c $(INCLDIR)/libo.h
	$(CC) $(COPTS) -o test-libo.obj -c $(SRCDIR)/test-libo.c

libo.a: libo.obj
	ar rcs libo.a libo.obj

libo.obj: $(SRCDIR)/libo.c $(INCLDIR)/libo.h
	$(CC) $(COPTS) -o libo.obj -c $(SRCDIR)/libo.c

o.lib: libo.a
	@cp libo.a o.lib
	ranlib o.lib

clean:
	@rm -f *.a *.obj *.lib *.exe

install: o.lib
	@cp o.lib $(WINPKGLIBDIR)
	@cp ../include/libo.h $(WINPKGINCLDIR)

uninstall:
	@rm -f $(WINPKGLIBDIR)/o.lib
	@rm -f $(WINPKGINCLDIR)/libo.h

