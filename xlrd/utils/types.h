#pragma once

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
        if (value == nullptr) {
            m_obj = nullptr;
            return *this;
        }
        m_obj = new _any<T>(value);
        return *this;
    }
    
    template<class T>
    const T& cast() const {
        return dynamic_cast< _any<T>& >(*m_obj).m_value;
    }
    
    const std::type_info& type () const {
        if (m_obj == nullptr) {
            return typeid(nullptr);
        }
        return m_obj->type();
    }
    
    bool is_null() {
        return m_obj == nullptr;
    }

    ~any () {
        delete m_obj;
    }
    
};

}
}
