#ifndef MYTRANDOM_HPP
#define MYTRANDOM_HPP

#include "MyTGL2/DataType/MytVector.hpp"

namespace Myt{
	template<class T_DataType>
	class Random{
	public:
		template<class T_RndProvider>
		static inline T_DataType Rnd(T_RndProvider &rnd){
			return (T_DataType)rnd.Rnd();
		}
	};

	template<>
	class Random<float>{
	public:
		template<class T_RndProvider>
		static inline float Rnd(T_RndProvider &rnd){
			return float(rnd.Rnd())*(1.0f/4294967296.0f);
		}
	};

	template<>
	class Random<double>{
	public:
		template<class T_RndProvider>
		static inline double Rnd(T_RndProvider &rnd){
			unsigned long a=rnd.Rnd()>>5, b=rnd.Rnd()>>6; 
			return(a*67108864.0+b)*(1.0/9007199254740992.0); 
		}
	};

	template<class T_DataType,int N_Size>
	class Random<Vector<T_DataType,N_Size> >{
	public:
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
	class Random<Vector<float,N_Size> >{
	public:
		template<class T_RndProvider>
		static inline Vector<float,N_Size> Rnd(T_RndProvider &rnd){
			Vector<float,N_Size> v;
			for(int i=0;i<N_Size;i++){
				v[i]=Random<float>::Rnd(rnd);
			}
			return v;
		}
	};
}

#endif