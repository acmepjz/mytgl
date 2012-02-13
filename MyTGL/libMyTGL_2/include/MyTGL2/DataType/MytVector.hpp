#ifndef MYTVECTOR_HPP
#define MYTVECTOR_HPP

#include "MyTGL2/Util/MytFunctions.hpp"

namespace Myt{
	template<class T_DataType,int N_Size>
	struct Vector{
	public:
		T_DataType Data[N_Size];
	public:
		typedef T_DataType DataType;

		inline int Size() const{
			return N_Size;
		}

		inline Vector& operator=(const T_DataType& Value){
			for(int i=0;i<N_Size;i++) Data[i]=Value;
			return *this;
		}

		inline Vector& operator+=(const T_DataType& Value){
			for(int i=0;i<N_Size;i++) Data[i]+=Value;
			return *this;
		}

		inline Vector& operator+=(const Vector& obj){
			for(int i=0;i<N_Size;i++) Data[i]+=obj.Data[i];
			return *this;
		}

		inline Vector& operator-=(const T_DataType& Value){
			for(int i=0;i<N_Size;i++) Data[i]-=Value;
			return *this;
		}

		inline Vector& operator-=(const Vector& obj){
			for(int i=0;i<N_Size;i++) Data[i]-=obj.Data[i];
			return *this;
		}

		inline Vector& operator*=(const T_DataType& Value){
			for(int i=0;i<N_Size;i++) Data[i]*=Value;
			return *this;
		}

		inline Vector& operator*=(const Vector& obj){
			for(int i=0;i<N_Size;i++) Data[i]*=obj.Data[i];
			return *this;
		}

		inline Vector& operator/=(const T_DataType& Value){
			for(int i=0;i<N_Size;i++) Data[i]/=Value;
			return *this;
		}

		inline Vector& operator/=(const Vector& obj){
			for(int i=0;i<N_Size;i++) Data[i]/=obj.Data[i];
			return *this;
		}

		inline friend Vector operator+(const Vector& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]+obj2.Data[i];
			return ret;
		}

		inline friend Vector operator+(const T_DataType& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1+obj2.Data[i];
			return ret;
		}

		inline friend Vector operator+(const Vector& obj1,T_DataType& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]+obj2;
			return ret;
		}

		inline friend Vector operator-(const Vector& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]-obj2.Data[i];
			return ret;
		}

		inline friend Vector operator-(const T_DataType& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1-obj2.Data[i];
			return ret;
		}

		inline friend Vector operator-(const Vector& obj1,const T_DataType& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]-obj2;
			return ret;
		}

		inline friend Vector operator*(const Vector& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]*obj2.Data[i];
			return ret;
		}

		inline friend Vector operator*(const T_DataType& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1*obj2.Data[i];
			return ret;
		}

		inline friend Vector operator*(const Vector& obj1,const T_DataType& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]*obj2;
			return ret;
		}

		inline friend Vector operator/(const Vector& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]/obj2.Data[i];
			return ret;
		}

		inline friend Vector operator/(const T_DataType& obj1,const Vector& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1/obj2.Data[i];
			return ret;
		}

		inline friend Vector operator/(const Vector& obj1,const T_DataType& obj2){
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=obj1.Data[i]/obj2;
			return ret;
		}

		inline Vector operator-() const{
			Vector ret;
			for(int i=0;i<N_Size;i++) ret.Data[i]=-Data[i];
			return ret;
		}

		inline const T_DataType& operator[](int idx) const{
			return Data[idx];
		}
		inline T_DataType& operator[](int idx){
			return Data[idx];
		}
	};

	template<class T_DataType,int N_Size>
	struct Functions<Vector<T_DataType,N_Size> >{
		static inline Vector<T_DataType,N_Size> Zero(){
			Vector<T_DataType,N_Size> v={};
			return v;
		}
		static inline Vector<T_DataType,N_Size> One(){
			Vector<T_DataType,N_Size> v;
			v=Functions<T_DataType>::One();
			return v;
		}
	};

	typedef Vector<unsigned char,1> Vector8UC1;
	typedef Vector<unsigned char,2> Vector8UC2;
	typedef Vector<unsigned char,3> Vector8UC3;
	typedef Vector<unsigned char,4> Vector8UC4;

	typedef Vector<float,1> Vector32FC1;
	typedef Vector<float,2> Vector32FC2;
	typedef Vector<float,3> Vector32FC3;
	typedef Vector<float,4> Vector32FC4;
}

#endif
