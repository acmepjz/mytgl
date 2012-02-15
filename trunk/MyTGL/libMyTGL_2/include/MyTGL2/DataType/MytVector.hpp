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
	struct Constants<Vector<T_DataType,N_Size> >{
		static inline Vector<T_DataType,N_Size> Zero(){
			Vector<T_DataType,N_Size> v={};
			return v;
		}
		static inline Vector<T_DataType,N_Size> One(){
			Vector<T_DataType,N_Size> v;
			v=Constants<T_DataType>::One();
			return v;
		}
	};

	template<class T_DataType,int N_Size,class T_float>
	struct LinearFunctions<Vector<T_DataType,N_Size>,T_float>{
		static inline Vector<T_DataType,N_Size> Lerp(const Vector<T_DataType,N_Size>& Value1,const Vector<T_DataType,N_Size>& Value2,T_float f){
			Vector<T_DataType,N_Size> v;
			for(int i=0;i<N_Size;i++) v[i]=T_DataType(T_float(Value1[i])+(T_float(Value2[i])-T_float(Value1[i]))*f);
			return v;
		}
	};

	template<class T_DataType,int N_Size>
	struct OrderFunctions<Vector<T_DataType,N_Size> >{
		static inline Vector<T_DataType,N_Size> Max(const Vector<T_DataType,N_Size>& Value1,const Vector<T_DataType,N_Size>& Value2){
			Vector<T_DataType,N_Size> v;
			for(int i=0;i<N_Size;i++) v[i]=Value1[i]>Value2[i]?Value1[i]:Value2[i];
			return v;
		}
		static inline Vector<T_DataType,N_Size> Min(const Vector<T_DataType,N_Size>& Value1,const Vector<T_DataType,N_Size>& Value2){
			Vector<T_DataType,N_Size> v;
			for(int i=0;i<N_Size;i++) v[i]=Value1[i]<Value2[i]?Value1[i]:Value2[i];
			return v;
		}
		static inline Vector<T_DataType,N_Size> Clamp(const Vector<T_DataType,N_Size>& Value,const Vector<T_DataType,N_Size>& Min,const Vector<T_DataType,N_Size>& Max){
			Vector<T_DataType,N_Size> v;
			for(int i=0;i<N_Size;i++) v[i]=Value[i]<Min[i]?Min[i]:(Value[i]>Max[i]?Max[i]:Value[i]);
			return v;
		}

		//Size should be >0
		static inline Vector<T_DataType,N_Size> FindMax(const Vector<T_DataType,N_Size>* Array,unsigned int Size){
			Vector<T_DataType,N_Size> t=Array[0];
			for(unsigned int idx=1;idx<Size;idx++){
				for(int i=0;i<N_Size;i++) if(Array[idx][i]>t[i]) t[i]=Array[idx][i];
			}
			return t;
		}
		//Size should be >0
		static inline Vector<T_DataType,N_Size> FindMin(const Vector<T_DataType,N_Size>* Array,unsigned int Size){
			Vector<T_DataType,N_Size> t=Array[0];
			for(unsigned int idx=1;idx<Size;idx++){
				for(int i=0;i<N_Size;i++) if(Array[idx][i]<t[i]) t[i]=Array[idx][i];
			}
			return t;
		}
		//Size should be >0
		static inline void FindMinMax(const Vector<T_DataType,N_Size>* Array,Vector<T_DataType,N_Size>& Min,Vector<T_DataType,N_Size>& Max,unsigned int Size){
			Min=Array[0];
			Max=Array[0];
			for(unsigned int idx=1;idx<Size;idx++){
				for(int i=0;i<N_Size;i++){
					if(Array[idx][i]<Min[i]) Min[i]=Array[idx][i];
					else if(Array[idx][i]>Max[i]) Max[i]=Array[idx][i];
				}
			}
		}
	};

	template<class T_DataType,int N_Size>
	struct Random<Vector<T_DataType,N_Size> >{
		template<class T_RndProvider>
		static inline Vector<T_DataType,N_Size> Rnd(T_RndProvider &rnd){
			if(sizeof(Vector<T_DataType,N_Size>)<=sizeof(unsigned long)){
				unsigned long a=rnd.Rnd();
				return (Vector<T_DataType,N_Size>&)a;
			}else{
				Vector<T_DataType,N_Size> v;
				for(int i=0;i<N_Size;i++){
					v[i]=Random<T_DataType>::Rnd(rnd);
				}
				return v;
			}
		}
	};

	template<int N_Size>
	struct Random<Vector<float,N_Size> >{
		template<class T_RndProvider>
		static inline Vector<float,N_Size> Rnd(T_RndProvider &rnd){
			Vector<float,N_Size> v;
			for(int i=0;i<N_Size;i++){
				v[i]=Random<float>::Rnd(rnd);
			}
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

	typedef Vector<double,1> Vector64FC1;
	typedef Vector<double,2> Vector64FC2;
	typedef Vector<double,3> Vector64FC3;
	typedef Vector<double,4> Vector64FC4;
}

#endif
