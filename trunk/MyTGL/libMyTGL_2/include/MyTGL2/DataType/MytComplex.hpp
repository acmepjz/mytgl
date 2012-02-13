#ifndef MYTCOMPLEX_HPP
#define MYTCOMPLEX_HPP

#include "MyTGL2/Util/MytFunctions.hpp"
#include <math.h>

namespace Myt{
	template<class T>
	struct Complex{
	public:
		T Re;
		T Im;
	public:
		typedef T DataType;

		static inline Complex Zero(){
			Complex c={};
			return c;
		}

		static inline Complex One(){
			Complex c={Functions<T>::One()};
			return c;
		}

		static inline Complex I(){
			Complex c={Functions<T>::Zero(),Functions<T>::One()};
			return c;
		}

		static inline Complex Make(const T& Re,const T& Im){
			Complex c={Re,Im};
			return c;
		}

		inline Complex& operator+=(const T& Value){
			Re+=Value;
			return *this;
		}

		inline Complex& operator+=(const Complex& Value){
			Re+=Value.Re;
			Im+=Value.Im;
			return *this;
		}

		inline Complex& operator-=(const T& Value){
			Re-=Value;
			return *this;
		}

		inline Complex& operator-=(const Complex& Value){
			Re-=Value.Re;
			Im-=Value.Im;
			return *this;
		}

		inline Complex& operator*=(const T& Value){
			Re*=Value;
			Im*=Value;
			return *this;
		}

		inline Complex& operator*=(const Complex& Value){
			T tmp=Re*Value.Im+Im*Value.Re;
			Re=Re*Value.Re-Im*Value.Im;
			Im=tmp;
			return *this;
		}

		inline Complex& operator/=(const T& Value){
			Re/=Value;
			Im/=Value;
			return *this;
		}

		inline Complex& operator/=(const Complex& Value){
			T norm=Value.Re*Value.Re+Value.Im*Value.Im;
			T tmp=(Im*Value.Re-Re*Value.Im)/norm;
			Re=(Re*Value.Re+Im*Value.Im)/norm;
			Im=tmp;
			return *this;
		}

		inline friend Complex operator+(const Complex& obj1,const Complex& obj2){
			Complex c={obj1.Re+obj2.Re,obj1.Im+obj2.Im};
			return c;
		}

		inline friend Complex operator+(const T& obj1,const Complex& obj2){
			Complex c={obj1+obj2.Re,obj2.Im};
			return c;
		}

		inline friend Complex operator+(const Complex& obj1,const T& obj2){
			Complex c={obj1.Re+obj2,obj1.Im};
			return c;
		}

		inline friend Complex operator-(const Complex& obj1,const Complex& obj2){
			Complex c={obj1.Re-obj2.Re,obj1.Im-obj2.Im};
			return c;
		}

		inline friend Complex operator-(const T& obj1,const Complex& obj2){
			Complex c={obj1-obj2.Re,-obj2.Im};
			return c;
		}

		inline friend Complex operator-(const Complex& obj1,const T& obj2){
			Complex c={obj1.Re-obj2,obj1.Im};
			return c;
		}

		inline friend Complex operator*(const Complex& obj1,const Complex& obj2){
			Complex c={obj1.Re*obj2.Re-obj1.Im*obj2.Im,obj1.Re*obj2.Im+obj1.Im*obj2.Re};
			return c;
		}

		inline friend Complex operator*(const T& obj1,const Complex& obj2){
			Complex c={obj1*obj2.Re,obj1*obj2.Im};
			return c;
		}

		inline friend Complex operator*(const Complex& obj1,const T& obj2){
			Complex c={obj1.Re*obj2,obj1.Im*obj2};
			return c;
		}

		inline friend Complex operator/(const Complex& obj1,const Complex& obj2){
			T tmp=obj2.Re*obj2.Re+obj2.Im*obj2.Im;
			Complex c={(obj1.Re*obj2.Re+obj1.Im*obj2.Im)/tmp,(obj1.Im*obj2.Re-obj1.Re*obj2.Im)/tmp};
			return c;
		}

		inline friend Complex operator/(const T& obj1,const Complex& obj2){
			T tmp=obj2.Re*obj2.Re+obj2.Im*obj2.Im;
			Complex c={obj1*obj2.Re/tmp,-obj1*obj2.Im/tmp};
			return c;
		}

		inline friend Complex operator/(const Complex& obj1,const T& obj2){
			Complex c={obj1.Re/obj2,obj1.Im/obj2};
			return c;
		}

		inline Complex operator-() const{
			Complex c={-Re,-Im};
			return c;
		}
	};

	template<class T>
	struct Functions<Complex<T> >{
		static inline Complex<T> Zero(){
			Complex<T> c={};
			return c;
		}
		static inline Complex<T> One(){
			Complex<T> c={Functions<T>::One()};
			return c;
		}
	};
}

#endif