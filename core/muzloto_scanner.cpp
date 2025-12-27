#include <string>
#include <vector>
#include <map>
#include <memory>
#include <opencv2/opencv.hpp>
#include <tesseract/baseapi.h>
#include <leptonica/allheaders.h>

#ifdef _WIN32
    #define MUZLOTO_EXPORT __declspec(dllexport)
#else
    #define MUZLOTO_EXPORT __attribute__((visibility("default")))
#endif

namespace muzloto {

struct FieldResult {
    std::string name;
    std::string value;
    float confidence;
};

struct ScanResult {
    bool success;
    std::string error_message;
    std::vector<FieldResult> fields;
    std::string raw_text;
    double processing_time_ms;
    
    // Конкретные поля анкеты Muzloto
    std::string date;
    std::string table_number;
    std::string location;
    std::string satisfaction;  // Довольны ли посещением
    std::string playlist_liked; // Понравился ли плейлист
    std::string tracks_to_add; // Какие треки добавить
    std::string location_liked; // Понравилась ли локация
    std::string kitchen_liked; // Понравились ли кухня и бар
    std::string service_ok;    // Устроил ли сервис
    std::string host_work;     // Понравилась ли работа ведущего
    std::string visits_count;  // Сколько раз были
    std::string ticket_price;  // Оценка стоимости
    std::string know_booking;  // Знают ли о заказе
    std::string source_info;   // Откуда узнали
    std::string purpose;       // Ради чего ходят
    std::string improvements;  // Что улучшить
    std::string phone_number;  // Телефон
};

class MUZLOTO_EXPORT MuzlotoScanner {
private:
    std::unique_ptr<tesseract::TessBaseAPI> ocr;
    bool initialized;
    
    // Шаблон полей анкеты Muzloto
    const std::vector<std::string> field_names = {
        "Дата",
        "Номер столика", 
        "Место игры",
        "Довольны ли вы посещением Музлого",
        "Понравился ли вам плейлист",
        "Какие треки вы бы добавили",
        "Понравилась ли вам локация",
        "Понравилась ли вам кухня и бар",
        "Устроил ли вас сервис, время подачи",
        "Понравилась ли вам работа ведущего",
        "Сколько раз вы были на Музлого",
        "Оцените стоимость игры за билет",
        "Знаете ли вы, что Музлого можно заказать",
        "Откуда вы о нас узнали",
        "Ради чего вы обычно ходите на подобные вечеринки",
        "Что нам стоит улучшить",
        "Телефон"
    };
    
public:
    MuzlotoScanner() : initialized(false) {
        ocr = std::make_unique<tesseract::TessBaseAPI>();
    }
    
    ~MuzlotoScanner() {
        if (ocr) {
            ocr->End();
        }
    }
    
    bool initialize(const std::string& tessdata_path = "") {
        try {
            // Инициализация Tesseract с русским языком
            if (ocr->Init(tessdata_path.empty() ? NULL : tessdata_path.c_str(), 
                         "rus+eng", 
                         tesseract::OEM_LSTM_ONLY) != 0) {
                return false;
            }
            
            // Настройки для анкет
            ocr->SetPageSegMode(tesseract::PSM_AUTO);
            ocr->SetVariable("preserve_interword_spaces", "1");
            ocr->SetVariable("textord_tabfind_find_tables", "1");
            ocr->SetVariable("textord_tablefind_recognize_tables", "1");
            
            initialized = true;
            return true;
            
        } catch (const std::exception& e) {
            std::cerr << "Ошибка инициализации: " << e.what() << std::endl;
            return false;
        }
    }
    
    ScanResult scan_image(const std::string& image_path) {
        ScanResult result;
        auto start_time = std::chrono::high_resolution_clock::now();
        
        try {
            if (!initialized) {
                throw std::runtime_error("Сканер не инициализирован");
            }
            
            // 1. Загрузка изображения
            cv::Mat image = cv::imread(image_path, cv::IMREAD_COLOR);
            if (image.empty()) {
                throw std::runtime_error("Не удалось загрузить изображение: " + image_path);
            }
            
            // 2. Предобработка
            cv::Mat processed = preprocess_image(image);
            
            // 3. Распознавание текста
            ocr->SetImage(processed.data, 
                         processed.cols, 
                         processed.rows, 
                         processed.channels(), 
                         processed.step);
            
            char* text = ocr->GetUTF8Text();
            result.raw_text = text ? std::string(text) : "";
            delete[] text;
            
            // 4. Парсинг анкеты Muzloto
            parse_muzloto_form(result);
            
            // 5. Извлечение ответов по шаблону
            extract_answers(result);
            
            result.success = true;
            
        } catch (const std::exception& e) {
            result.success = false;
            result.error_message = e.what();
        }
        
        auto end_time = std::chrono::high_resolution_clock::now();
        result.processing_time_ms = std::chrono::duration<double, std::milli>(
            end_time - start_time).count();
        
        return result;
    }
    
private:
    cv::Mat preprocess_image(const cv::Mat& image) {
        cv::Mat gray, denoised, binary;
        
        // Конвертация в оттенки серого
        cv::cvtColor(image, gray, cv::COLOR_BGR2GRAY);
        
        // Удаление шума
        cv::fastNlMeansDenoising(gray, denoised, 10, 7, 21);
        
        // Улучшение контраста
        cv::Mat equalized;
        cv::equalizeHist(denoised, equalized);
        
        // Адаптивная бинаризация
        cv::adaptiveThreshold(equalized, binary, 255,
                             cv::ADAPTIVE_THRESH_GAUSSIAN_C,
                             cv::THRESH_BINARY, 11, 2);
        
        return binary;
    }
    
    void parse_muzloto_form(ScanResult& result) {
        // Разбиваем текст на строки
        std::vector<std::string> lines;
        std::stringstream ss(result.raw_text);
        std::string line;
        
        while (std::getline(ss, line)) {
            if (!line.empty()) {
                lines.push_back(line);
            }
        }
        
        // Ищем поля анкеты в тексте
        for (const auto& field_name : field_names) {
            for (size_t i = 0; i < lines.size(); i++) {
                if (lines[i].find(field_name) != std::string::npos) {
                    FieldResult field;
                    field.name = field_name;
                    
                    // Берем следующую строку как ответ
                    if (i + 1 < lines.size()) {
                        field.value = lines[i + 1];
                    }
                    
                    field.confidence = 0.9f; // Предполагаемая уверенность
                    result.fields.push_back(field);
                    break;
                }
            }
        }
        
        // Извлекаем конкретные поля для удобства
        for (const auto& field : result.fields) {
            if (field.name.find("Дата") != std::string::npos) {
                result.date = field.value;
            } else if (field.name.find("Номер столика") != std::string::npos) {
                result.table_number = field.value;
            } else if (field.name.find("Место игры") != std::string::npos) {
                result.location = field.value;
            } else if (field.name.find("Довольны") != std::string::npos) {
                result.satisfaction = field.value;
            } else if (field.name.find("плейлист") != std::string::npos || 
                      field.name.find("пленить") != std::string::npos) {
                result.playlist_liked = field.value;
            } else if (field.name.find("треки") != std::string::npos) {
                result.tracks_to_add = field.value;
            } else if (field.name.find("локация") != std::string::npos) {
                result.location_liked = field.value;
            } else if (field.name.find("кухня") != std::string::npos) {
                result.kitchen_liked = field.value;
            } else if (field.name.find("сервис") != std::string::npos) {
                result.service_ok = field.value;
            } else if (field.name.find("ведущего") != std::string::npos) {
                result.host_work = field.value;
            } else if (field.name.find("Сколько раз") != std::string::npos) {
                result.visits_count = field.value;
            } else if (field.name.find("стоимость") != std::string::npos) {
                result.ticket_price = field.value;
            } else if (field.name.find("заказать") != std::string::npos) {
                result.know_booking = field.value;
            } else if (field.name.find("Откуда") != std::string::npos) {
                result.source_info = field.value;
            } else if (field.name.find("Ради чего") != std::string::npos) {
                result.purpose = field.value;
            } else if (field.name.find("улучшить") != std::string::npos) {
                result.improvements = field.value;
            } else if (field.name.find("Телефон") != std::string::npos || 
                      field.name.find("номер") != std::string::npos) {
                result.phone_number = extract_phone_number(field.value);
            }
        }
    }
    
    void extract_answers(ScanResult& result) {
        // Дополнительная обработка ответов
        if (!result.phone_number.empty()) {
            // Нормализация телефона
            result.phone_number = normalize_phone(result.phone_number);
        }
        
        // Для полей с галочками/отметками
        auto is_checked = [](const std::string& text) -> std::string {
            if (text.find("да") != std::string::npos || 
                text.find("Да") != std::string::npos ||
                text.find("ДА") != std::string::npos ||
                text.find("✓") != std::string::npos ||
                text.find("+") != std::string::npos ||
                text.find("V") != std::string::npos) {
                return "Да";
            } else if (text.find("нет") != std::string::npos ||
                      text.find("Нет") != std::string::npos ||
                      text.find("НЕТ") != std::string::npos) {
                return "Нет";
            }
            return text;
        };
        
        // Применяем к полям с да/нет
        result.satisfaction = is_checked(result.satisfaction);
        result.playlist_liked = is_checked(result.playlist_liked);
        result.location_liked = is_checked(result.location_liked);
        result.kitchen_liked = is_checked(result.kitchen_liked);
        result.service_ok = is_checked(result.service_ok);
        result.host_work = is_checked(result.host_work);
        result.know_booking = is_checked(result.know_booking);
    }
    
    std::string extract_phone_number(const std::string& text) {
        // Простая экстракция телефона
        std::regex phone_regex(R"((\+7|8)[\s\-\(]?(\d{3})[\s\-\)]?(\d{3})[\s\-]?(\d{2})[\s\-]?(\d{2}))");
        std::smatch match;
        
        if (std::regex_search(text, match, phone_regex)) {
            return match.str();
        }
        
        return "";
    }
    
    std::string normalize_phone(const std::string& phone) {
        std::string normalized = phone;
        
        // Убираем все нецифровые символы кроме +
        normalized.erase(std::remove_if(normalized.begin(), normalized.end(),
                         [](char c) { return !std::isdigit(c) && c != '+'; }),
                         normalized.end());
        
        // Приводим к формату +7XXXXXXXXXX
        if (!normalized.empty()) {
            if (normalized[0] == '8') {
                normalized[0] = '7';
                normalized = "+" + normalized;
            } else if (normalized.substr(0, 2) != "+7") {
                normalized = "+7" + normalized;
            }
        }
        
        return normalized;
    }
};

// C-интерфейс для простого использования
extern "C" {
    MUZLOTO_EXPORT void* muzloto_create() {
        return new MuzlotoScanner();
    }
    
    MUZLOTO_EXPORT void muzloto_destroy(void* scanner) {
        delete static_cast<MuzlotoScanner*>(scanner);
    }
    
    MUZLOTO_EXPORT int muzloto_initialize(void* scanner, const char* tessdata_path) {
        return static_cast<MuzlotoScanner*>(scanner)->initialize(
            tessdata_path ? std::string(tessdata_path) : ""
        ) ? 1 : 0;
    }
    
    MUZLOTO_EXPORT const char* muzloto_scan_image(void* scanner, const char* image_path) {
        auto result = static_cast<MuzlotoScanner*>(scanner)->scan_image(image_path);
        
        // Конвертируем результат в JSON
        nlohmann::json j;
        j["success"] = result.success;
        j["error_message"] = result.error_message;
        j["processing_time_ms"] = result.processing_time_ms;
        
        // Поля анкеты
        j["date"] = result.date;
        j["table_number"] = result.table_number;
        j["location"] = result.location;
        j["satisfaction"] = result.satisfaction;
        j["playlist_liked"] = result.playlist_liked;
        j["tracks_to_add"] = result.tracks_to_add;
        j["location_liked"] = result.location_liked;
        j["kitchen_liked"] = result.kitchen_liked;
        j["service_ok"] = result.service_ok;
        j["host_work"] = result.host_work;
        j["visits_count"] = result.visits_count;
        j["ticket_price"] = result.ticket_price;
        j["know_booking"] = result.know_booking;
        j["source_info"] = result.source_info;
        j["purpose"] = result.purpose;
        j["improvements"] = result.improvements;
        j["phone_number"] = result.phone_number;
        
        j["raw_text"] = result.raw_text.substr(0, 500); // Первые 500 символов
        
        // Все распознанные поля
        nlohmann::json fields_array = nlohmann::json::array();
        for (const auto& field : result.fields) {
            nlohmann::json f;
            f["name"] = field.name;
            f["value"] = field.value;
            f["confidence"] = field.confidence;
            fields_array.push_back(f);
        }
        j["fields"] = fields_array;
        
        std::string json_str = j.dump();
        char* c_str = new char[json_str.length() + 1];
        std::strcpy(c_str, json_str.c_str());
        
        return c_str;
    }
    
    MUZLOTO_EXPORT void muzloto_free_string(const char* str) {
        delete[] str;
    }
}

} // namespace muzloto