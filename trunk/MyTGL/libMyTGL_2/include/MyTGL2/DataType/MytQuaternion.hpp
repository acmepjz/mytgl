#ifndef MYTQUATERNION_HPP
#define MYTQUATERNION_HPP

#include "MyTGL2/Util/MytFunctions.hpp"
#include <math.h>

namespace Myt{
	template<class T>
	struct Quaternion{
	public:
		T Data[4];
	public:
		typedef T DataType;

		static inline Quaternion Zero(){
			Quaternion c={};
			return c;
		}

		static inline Quaternion One(){
			Quaternion c={Functions<T>::One()};
			return c;
		}

		static inline Quaternion I(){
			Quaternion c={Functions<T>::Zero(),Functions<T>::One()};
			return c;
		}

		static inline Quaternion J(){
			Quaternion c={Functions<T>::Zero(),Functions<T>::Zero(),Functions<T>::One()};
			return c;
		}

		static inline Quaternion K(){
			Quaternion c={Functions<T>::Zero(),Functions<T>::Zero(),Functions<T>::Zero(),Functions<T>::One()};
			return c;
		}

		static inline Quaternion Make(const T& Data0,const T& Data1,const T& Data2,const T& Data3){
			Quaternion c={Data0,Data1,Data2,Data3};
			return c;
		}

		inline Quaternion& operator+=(const T& Value){
			Data[0]+=Value;
			return *this;
		}

		inline Quaternion& operator+=(const Quaternion& Value){
			for(int i=0;i<4;i++) Data[i]+=Value.Data[i];
			return *this;
		}

		inline Quaternion& operator-=(const T& Value){
			Data[0]-=Value;
			return *this;
		}

		inline Quaternion& operator-=(const Quaternion& Value){
			for(int i=0;i<4;i++) Data[i]-=Value.Data[i];
			return *this;
		}

		inline Quaternion& operator*=(const T& Value){
			for(int i=0;i<4;i++) Data[i]*=Value;
			return *this;
		}

		inline Quaternion& operator*=(const Quaternion& Value){
			T tmp0=Data[0]*Value.Data[0]-Data[1]*Value.Data[1]-Data[2]*Value.Data[2]-Data[3]*Value.Data[3];
			T tmp1=Data[1]*Value.Data[0]+Data[0]*Value.Data[1]+Data[2]*Value.Data[3]-Data[3]*Value.Data[2];
			T tmp2=Data[2]*Value.Data[0]+Data[0]*Value.Data[2]+Data[3]*Value.Data[1]-Data[1]*Value.Data[3];
			Data[3]=Data[3]*Value.Data[0]+Data[0]*Value.Data[3]+Data[1]*Value.Data[2]-Data[2]*Value.Data[1];
			Data[0]=tmp0;
			Data[1]=tmp1;
			Data[2]=tmp2;
			return *this;
		}

		inline Quaternion& operator/=(const T& Value){
			for(int i=0;i<4;i++) Data[i]/=Value;
			return *this;
		}

		inline Quaternion& operator/=(const Quaternion& Value){
			T tmp=Value.Data[0]*Value.Data[0]+Value.Data[1]*Value.Data[1]+Value.Data[2]*Value.Data[2]+Value.Data[3]*Value.Data[3];
			T tmp0=(Data[0]*Value.Data[0]+Data[1]*Value.Data[1]+Data[2]*Value.Data[2]+Data[3]*Value.Data[3])/tmp;
			T tmp1=(Data[1]*Value.Data[0]-Data[0]*Value.Data[1]-Data[2]*Value.Data[3]+Data[3]*Value.Data[2])/tmp;
			T tmp2=(Data[2]*Value.Data[0]-Data[0]*Value.Data[2]-Data[3]*Value.Data[1]+Data[1]*Value.Data[3])/tmp;
			Data[3]=(Data[3]*Value.Data[0]-Data[0]*Value.Data[3]-Data[1]*Value.Data[2]+Data[2]*Value.Data[1])/tmp;
			Data[0]=tmp0;
			Data[1]=tmp1;
			Data[2]=tmp2;
			return *this;
		}

		inline friend Quaternion operator+(const Quaternion& obj1,const Quaternion& obj2){
			Quaternion c={obj1.Data[0]+obj2.Data[0],obj1.Data[1]+obj2.Data[1],obj1.Data[2]+obj2.Data[2],obj1.Data[3]+obj2.Data[3]};
			return c;
		}

		inline friend Quaternion operator+(const T& obj1,const Quaternion& obj2){
			Quaternion c={obj1+obj2.Data[0],obj2.Data[1],obj2.Data[2],obj2.Data[3]};
			return c;
		}

		inline friend Quaternion operator+(const Quaternion& obj1,const T& obj2){
			Quaternion c={obj1.Data[0]+obj2,obj1.Data[1],obj1.Data[2],obj1.Data[3]};
			return c;
		}

		inline friend Quaternion operator-(const Quaternion& obj1,const Quaternion& obj2){
			Quaternion c={obj1.Data[0]-obj2.Data[0],obj1.Data[1]-obj2.Data[1],obj1.Data[2]-obj2.Data[2],obj1.Data[3]-obj2.Data[3]};
			return c;
		}

		inline friend Quaternion operator-(const T& obj1,const Quaternion& obj2){
			Quaternion c={obj1-obj2.Data[0],-obj2.Data[1],-obj2.Data[2],-obj2.Data[3]};
			return c;
		}

		inline friend Quaternion operator-(const Quaternion& obj1,const T& obj2){
			Quaternion c={obj1.Data[0]-obj2,obj1.Data[1],obj1.Data[2],obj1.Data[3]};
			return c;
		}

		inline friend Quaternion operator*(const Quaternion& obj1,const Quaternion& obj2){
			Quaternion c={
				obj1.Data[0]*obj2.Data[0]-obj1.Data[1]*obj2.Data[1]-obj1.Data[2]*obj2.Data[2]-obj1.Data[3]*obj2.Data[3],
				obj1.Data[1]*obj2.Data[0]+obj1.Data[0]*obj2.Data[1]+obj1.Data[2]*obj2.Data[3]-obj1.Data[3]*obj2.Data[2],
				obj1.Data[2]*obj2.Data[0]+obj1.Data[0]*obj2.Data[2]+obj1.Data[3]*obj2.Data[1]-obj1.Data[1]*obj2.Data[3],
				obj1.Data[3]*obj2.Data[0]+obj1.Data[0]*obj2.Data[3]+obj1.Data[1]*obj2.Data[2]-obj1.Data[2]*obj2.Data[1]
			};
			return c;
		}

		inline friend Quaternion operator*(const T& obj1,const Quaternion& obj2){
			Quaternion c={obj1*obj2.Data[0],obj1*obj2.Data[1],obj1*obj2.Data[2],obj1*obj2.Data[3]};
			return c;
		}

		inline friend Quaternion operator*(const Quaternion& obj1,const T& obj2){
			Quaternion c={obj1.Data[0]*obj2,obj1.Data[1]*obj2,obj1.Data[2]*obj2,obj1.Data[3]*obj2};
			return c;
		}

		//NOTE: Quaternion a/b is not well-defined; it can be a*b^{-1} or b^{-1}*a
		//this function calculates a*b^{-1}

		inline friend Quaternion operator/(const Quaternion& obj1,const Quaternion& obj2){
			T tmp=obj2.Data[0]*obj2.Data[0]+obj2.Data[1]*obj2.Data[1]+obj2.Data[2]*obj2.Data[2]+obj2.Data[3]*obj2.Data[3];
			Quaternion c={
				(obj1.Data[0]*obj2.Data[0]+obj1.Data[1]*obj2.Data[1]+obj1.Data[2]*obj2.Data[2]+obj1.Data[3]*obj2.Data[3])/tmp,
				(obj1.Data[1]*obj2.Data[0]-obj1.Data[0]*obj2.Data[1]-obj1.Data[2]*obj2.Data[3]+obj1.Data[3]*obj2.Data[2])/tmp,
				(obj1.Data[2]*obj2.Data[0]-obj1.Data[0]*obj2.Data[2]-obj1.Data[3]*obj2.Data[1]+obj1.Data[1]*obj2.Data[3])/tmp,
				(obj1.Data[3]*obj2.Data[0]-obj1.Data[0]*obj2.Data[3]-obj1.Data[1]*obj2.Data[2]+obj1.Data[2]*obj2.Data[1])/tmp
			};
			return c;
		}

		inline friend Quaternion operator/(const T& obj1,const Quaternion& obj2){
			T tmp=obj2.Data[0]*obj2.Data[0]+obj2.Data[1]*obj2.Data[1]+obj2.Data[2]*obj2.Data[2]+obj2.Data[3]*obj2.Data[3];
			Quaternion c={obj1*obj2.Data[0]/tmp,-obj1*obj2.Data[1]/tmp,-obj1*obj2.Data[2]/tmp,-obj1*obj2.Data[3]/tmp};
			return c;
		}

		inline friend Quaternion operator/(const Quaternion& obj1,const T& obj2){
			Quaternion c={obj1.Data[0]/obj2,obj1.Data[1]/obj2,obj1.Data[2]/obj2,obj1.Data[3]/obj2};
			return c;
		}

		inline Quaternion operator-() const{
			Quaternion c={-Data[0],-Data[1],-Data[2],-Data[3]};
			return c;
		}
	};

	template<class T>
	struct Functions<Quaternion<T> >{
		static inline Quaternion<T> Zero(){
			Quaternion<T> c={};
			return c;
		}
		static inline Quaternion<T> One(){
			Quaternion<T> c={Functions<T>::One()};
			return c;
		}
	};
}

#endif