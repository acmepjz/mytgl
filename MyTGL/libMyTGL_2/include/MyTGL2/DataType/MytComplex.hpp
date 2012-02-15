#ifndef MYTCOMPLEX_HPP
#define MYTCOMPLEX_HPP

#include "MyTGL2/Util/MytFunctions.hpp"
#include "MyTGL2/DataType/MytVector.hpp"
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
			Complex c={Constants<T>::One()};
			return c;
		}

		static inline Complex I(){
			Complex c={Constants<T>::Zero(),Constants<T>::One()};
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

		template<class T2>
		inline friend Complex operator*(const Complex& obj1,const Complex<T2>& obj2){
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

		template<class T2>
		inline friend Complex operator/(const Complex& obj1,const Complex<T2>& obj2){
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

		//conjugate
		inline Complex operator!() const{
			Complex c={Re,-Im};
			return c;
		}
	};

	template<class T>
	struct Constants<Complex<T> >{
		static inline Complex<T> Zero(){
			Complex<T> c={};
			return c;
		}
		static inline Complex<T> One(){
			Complex<T> c={Constants<T>::One()};
			return c;
		}
	};

	template<class T>
	struct Random<Complex<T> >{
		template<class T_RndProvider>
		static inline Complex<T> Rnd(T_RndProvider &rnd){
			Complex<T> c={Random<T>::Rnd(rnd),Random<T>::Rnd(rnd)};
			return c;
		}
	};

	typedef Complex<float> Complex32F;
	typedef Complex<double> Complex64F;

	typedef Complex<Vector32FC1> Complex32FC1;
	typedef Complex<Vector32FC2> Complex32FC2;
	typedef Complex<Vector32FC3> Complex32FC3;
	typedef Complex<Vector32FC4> Complex32FC4;

	typedef Complex<Vector64FC1> Complex64FC1;
	typedef Complex<Vector64FC2> Complex64FC2;
	typedef Complex<Vector64FC3> Complex64FC3;
	typedef Complex<Vector64FC4> Complex64FC4;
}

#endif