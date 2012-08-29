#ifndef MYTFUNCTIONS_HPP
#define MYTFUNCTIONS_HPP

#include <math.h>

namespace Myt{
	template<class T>
	struct Constants{
		static inline T Zero(){
			return 0;
		}
		static inline T One(){
			return 1;
		}
	};

	template<class T,class T_float>
	struct LinearFunctions{
		static inline T Lerp(const T& Value1,const T& Value2,T_float f){
			return T(T_float(Value1)+(T_float(Value2)-T_float(Value1))*f);
		}
	};

	template<class T>
	struct Random{
		template<class T_RndProvider>
		static inline T Rnd(T_RndProvider &rnd){
			return (T)rnd.Rnd();
		}
	};

	template<class T>
	struct OrderFunctions{
		static inline T Max(const T& Value1,const T& Value2){
			return Value1>Value2?Value1:Value2;
		}
		static inline T Min(const T& Value1,const T& Value2){
			return Value1<Value2?Value1:Value2;
		}
		static inline T Clamp(const T& Value,const T& Min,const T& Max){
			return Value<Min?Min:(Value>Max?Max:Value);
		}

		//Size should be >0
		static inline T FindMax(const T* Array,unsigned int Size){
			T t=Array[0];
			for(unsigned int idx=1;idx<Size;idx++){
				if(Array[idx]>t) t=Array[idx];
			}
			return t;
		}
		//Size should be >0
		static inline T FindMin(const T* Array,unsigned int Size){
			T t=Array[0];
			for(unsigned int idx=1;idx<Size;idx++){
				if(Array[idx]<t) t=Array[idx];
			}
			return t;
		}
		//Size should be >0
		static inline void FindMinMax(const T* Array,T& Min,T& Max,unsigned int Size){
			Min=Array[0];
			Max=Array[0];
			for(unsigned int idx=1;idx<Size;idx++){
				if(Array[idx]<Min) Min=Array[idx];
				else if(Array[idx]>Max) Max=Array[idx];
			}
		}
	};

	template<class T>
	struct LinearOrderFunctions{
		static int QSortCompareFunction(const void* lp1_,const void* lp2_){
			const T& lp1=*(const T*)lp1_;
			const T& lp2=*(const T*)lp2_;
			return lp1<lp2?-1:(lp1>lp2?1:0);
		}
	};

	template<>
	struct Random<float>{
		template<class T_RndProvider>
		static inline float Rnd(T_RndProvider &rnd){
			return float(rnd.Rnd())*(1.0f/4294967296.0f);
		}
	};

	template<>
	struct Random<double>{
		template<class T_RndProvider>
		static inline double Rnd(T_RndProvider &rnd){
			unsigned long a=rnd.Rnd()>>5, b=rnd.Rnd()>>6; 
			return(a*67108864.0+b)*(1.0/9007199254740992.0); 
		}
	};

	class IRandom{
	public:
		virtual unsigned int Rnd()=0; 
	};
	template<class T_RndProvider>
	class IRandomImpl:public IRandom,public T_RndProvider{
	public:
		virtual unsigned int Rnd(){
			return static_cast<T_RndProvider*>(this)->Rnd();
		}
	};

	class INoise{
	public:
		virtual int Noise1(int x,int Seed)=0;
		virtual int Noise2(int x,int y,int Seed)=0;
		virtual int Noise3(int x,int y,int z,int Seed)=0;
	};
	template<class T_NoiseProvider>
	class INoiseImpl:public INoise,public T_NoiseProvider{
	public:
		virtual int Noise1(int x,int Seed){
			return static_cast<T_NoiseProvider*>(this)->Noise1(x,Seed);
		}
		virtual int Noise2(int x,int y,int Seed){
			return static_cast<T_NoiseProvider*>(this)->Noise2(x,y,Seed);
		}
		virtual int Noise3(int x,int y,int z,int Seed){
			return static_cast<T_NoiseProvider*>(this)->Noise3(x,y,z,Seed);
		}
	};
}

#endif