PROJECT (libexcel C)
CMAKE_MINIMUM_REQUIRED (VERSION 2.8)

SET(CMAKE_C_FLAGS "-Wall -O2 -pipe")
INCLUDE_DIRECTORIES(include)

LIST(APPEND libexcel_src src/format.c src/hashhelp.c src/stream.c src/worksheet.c src/biffwriter.c src/formula.c src/olewriter.c src/workbook.c src/io_handler.c)

ADD_LIBRARY(excel STATIC ${libexcel_src})

ADD_SUBDIRECTORY(tests)
