cmake_minimum_required(VERSION 2.6)

project (MiniExcelReader)


set(tinyxml2_src 
    tinyxml2/tinyxml2.cpp
)

set(minizip_src
    minizip/ioapi.c
    minizip/ioapi_mem.c
    minizip/unzip.c
)


set(zlib_src
    zlib/adler32.c
    zlib/compress.c
    zlib/crc32.c
    zlib/deflate.c
    zlib/gzclose.c
    zlib/gzlib.c
    zlib/gzread.c
    zlib/gzwrite.c
    zlib/inflate.c
    zlib/infback.c
    zlib/inftrees.c
    zlib/inffast.c
    zlib/trees.c
    zlib/uncompr.c
    zlib/zutil.c
)

file(GLOB MiniExcelReader_src *.cpp)

set(libs tinyxml2 zlib minizip)

foreach(lib ${libs})
    foreach(src_file ${${lib}_src})
        set(${src_file} )
        get_filename_component (PATH ${src_file} PATH)
        string (REPLACE ${CMAKE_CURRENT_SOURCE_DIR} "" PATH ${PATH})
        string (REPLACE / \\ PATH ${PATH})
        source_group ("Source Files\\${PATH}" FILES ${src_file})
    endforeach()

    list(APPEND libsrc ${${lib}_src})
endforeach()

include_directories(${libs})
add_definitions(-D_CRT_SECURE_NO_WARNINGS)

add_executable(MiniExcelReader ${MiniExcelReader_src} ${libsrc})

