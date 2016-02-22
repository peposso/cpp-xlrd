#pragma once

#include <string>
#include <typeinfo>

namespace utils {
namespace types {

class any {
private:
    // 非テンプレート基本クラス
    struct _any_base {
        virtual ~_any_base() {}
        virtual const std::type_info& type () const = 0;
        virtual _any_base* clone() const = 0;
    };
    
    // テンプレート派生クラス
    template<class T>
    struct _any : public _any_base {
        T m_value;
        
        _any(T value) {
            m_value = value;
        }
        
        const std::type_info& type () const {
            return typeid(T);
        }
        
        _any_base* clone () const {
            return new _any<T>(m_value);
        }
        
        virtual ~_any() {}
    };
    
    _any_base* m_obj;

public:
    any () {
        m_obj = nullptr;
    }
    
    template<class T>
    any (const T& value) {
        m_obj = new _any<T>(value);
    }

    any (const char* value) {
        m_obj = new _any<std::string>(std::string(value));
    }
    
    any (const any& obj) {
        if ( obj.m_obj ) {
            m_obj = obj.m_obj->clone();
        }
        else {
            m_obj = 0;
        }
    }
    
    any& operator=(const any& obj) {
        delete m_obj;
        if ( obj.m_obj ) {
            m_obj = obj.m_obj->clone();
        }
        else {
            m_obj = 0;
        }
        return *this;
    }
    
    template<class T>
    any& operator=(const T& value) {
        delete m_obj;
        m_obj = new _any<T>(value);
        return *this;
    }

    any& operator=(std::nullptr_t value) {
        delete m_obj;
        if (value == nullptr) { m_obj = nullptr; }
        return *this;
    }
    
    template<class T>
    const T& cast() const {
        return dynamic_cast< _any<T>& >(*m_obj).m_value;
    }

    template<class T>
    const T& unsafe_cast() const {
        return static_cast< _any<T>& >(*m_obj).m_value;
    }

    template<class T>
    const bool is() const {
        return typeid(T) == m_obj->type();
    }

    const std::type_info& type () const {
        if (m_obj == nullptr) {
            return typeid(nullptr);
        }
        return m_obj->type();
    }
    
    const bool is_null() const {
        return m_obj == nullptr;
    }

    const bool is_str() const {
        return typeid(std::string) == m_obj->type();
    }

    const bool is_int() const {
        auto& type = m_obj->type();
        return (
            typeid(int) == type ||
            typeid(uint32_t) == type ||
            // typeid(int64_t) == type ||
            // typeid(uint64_t) == type ||
            typeid(int16_t) == type ||
            typeid(uint16_t) == type ||
            typeid(int8_t) == type ||
            typeid(uint8_t) == type
        );
    }

    const bool is_double() const {
        auto& type = m_obj->type();
        return (
            typeid(double) == type ||
            typeid(float) == type
        );
    }

    const int to_int() const {
        if (this->is_int()) {
            return cast<int>();
        }
        if (this->is_double()) {
            return int(cast<int>());
        }
        if (this->is<bool>()) {
            return cast<bool>() ? 1: 0;
        }
        if (this->is_str()) {
            return std::stoi(cast<std::string>());
        }
        return 0;
    }

    const double to_double() const {
        if (this->is_int()) {
            return double(cast<int>());
        }
        if (this->is_double()) {
            return cast<double>();
        }
        if (this->is<bool>()) {
            return cast<bool>() ? 1.0: 0.0;
        }
        if (this->is_str()) {
            return std::stod(cast<std::string>());
        }
        return 0.0;
    }

    const std::string to_str() const {
        if (this->is_str()) {
            return cast<std::string>();
        }
        if (this->is_int()) {
            return std::to_string(to_int());
        }
        if (this->is_double()) {
            return std::to_string(to_double());
        }
        if (this->is<bool>()) {
            return cast<bool>()? "true": "false";
        }
        return "";
    }

    ~any () {
        delete m_obj;
    }
    
};

}
}
