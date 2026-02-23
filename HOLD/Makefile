CC = gcc
COPTS = -O3 -g0 -Wall

all: libo.a test-libo doxygen

doxygen: Doxygen.libo libo.h libo.c README.md
	doxygen Doxygen.libo
	@touch doxygen

test-libo: test-libo.o libo.a
	$(CC) $(COPTS) -L. -o test-libo test-libo.o -lo -lzip $$(pkg-config --libs libxml-2.0) $$(pkg-config --libs libzip) -lstrings -lavl

test-libo.o: test-libo.c libo.h
	$(CC) $(COPTS) $$(pkg-config --cflags libxml-2.0 libzip) -c test-libo.c

libo.a: libo.o
	@(rm -f libo.a)
	ar r libo.a libo.o

libo.o: libo.c libo.h
	$(CC) $(COPTS) $$(pkg-config --cflags libxml-2.0) $$(pkg-config --cflags libzip) -c libo.c

install: libo.a
	@(mkdir -p /usr/local/lib)
	@(mkdir -p /usr/local/include)
	@(cp libo.a /usr/local/lib)
	@(cp libo.h /usr/local/include)

tar: clean
	@(cd ..; tar cJf libo.txz libo; cd -)

zip: clean
	@(cd ..; \
		rm -f libo.zip; \
		zip -r libo.zip libo; \
	)

clean:
	@(rm -f *.a *.o)
	@(rm -f *.obj *.exe *.dll *.lib)
	@(rm -f *.txz *.zip)
	@(rm -rf doxygen)
	@(rm -f test-libo)
	@(rm -f diffs)
	@(rm -rf explode mod-explode)
	@(rm -f TEST-CREATION.xlsx mod-TEST-CREATION.xlsx)

windows: clean
	@make -f Makefile.win

windows-install: windows
	@(cp o.lib $(HOME)/develop/win-pkgs/lib)
	@(cp libo.h $(HOME)/develop/win-pkgs/include)

windows-uninstall:
	@(rm -f  $(HOME)/develop/win-pkgs/lib/o.lib)
	@(rm -f  $(HOME)/develop/win-pkgs/include/libo.h)

diffs: exploded mod-exploded format-xml
	@( \
    rm -f diffs; \
    for f in $$(find explode -name '*.xml'); \
    do \
      echo FILE: $$(basename $$f) >> diffs; \
      diff "$$f" "mod-$$f" >> diffs; \
    done; \
    exit 0; \
   )

exploded:
	@(rm -rf explode; mkdir explode; cd explode; unzip -q ../TEST-CREATION.xlsx)

mod-exploded:
	@(rm -rf mod-explode; mkdir mod-explode; cd mod-explode; unzip -q ../mod-TEST-CREATION.xlsx)

format-xml:
	@(for f in $$(find explode -name '*.xml'); \
    do \
      xmllint --format "$$f" > /tmp/$$$$.xml; \
      mv /tmp/$$$$.xml "$$f"; \
    done)
	@(for f in $$(find mod-explode -name '*.xml'); \
    do \
      xmllint --format "$$f" > /tmp/$$$$.xml; \
      mv /tmp/$$$$.xml "$$f"; \
    done)

