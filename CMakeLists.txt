cmake_minimum_required(VERSION 3.15)
project(muzloto_scanner VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Поиск OpenCV
find_package(OpenCV REQUIRED)

# Поиск Tesseract
find_package(PkgConfig)
pkg_check_modules(TESSERACT REQUIRED tesseract lept)

# Настройка для разных ОС
if(WIN32)
    add_definitions(-DNOMINMAX -D_USE_MATH_DEFINES)
    set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS ON)
elseif(APPLE)
    set(CMAKE_MACOSX_RPATH ON)
else()
    # Linux
    add_compile_options(-fPIC)
endif()

# Создание библиотеки
add_library(muzloto_core SHARED
    core/muzloto_scanner.cpp
)

# Подключение зависимостей
target_include_directories(muzloto_core PRIVATE
    ${OpenCV_INCLUDE_DIRS}
    ${TESSERACT_INCLUDE_DIRS}
)

target_link_libraries(muzloto_core PRIVATE
    ${OpenCV_LIBS}
    ${TESSERACT_LIBRARIES}
)

# Установка
install(TARGETS muzloto_core
    LIBRARY DESTINATION lib
    RUNTIME DESTINATION bin
    ARCHIVE DESTINATION lib
)

