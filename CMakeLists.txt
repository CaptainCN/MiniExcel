cmake_minimum_required(VERSION 2.6)

project (MiniExcel)


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

set(MiniExcel_src MiniExcelReader.cpp)
set(MiniExcelLua_src LuaMiniExcel.cpp LuaMiniExcel.def)

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

add_executable(MiniExcelTest test.cpp ${MiniExcel_src} ${libsrc})

if (BUILD_LUA)
    find_package(Lua51 REQUIRED)
    include_directories(${LUA_INCLUDE_DIR} ${libs})
    add_library(miniexcel SHARED ${MiniExcel_src} ${MiniExcelLua_src} ${libsrc})
    target_link_libraries(miniexcel ${LUA_LIBRARIES})  
endif()
