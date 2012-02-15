#ifndef MYTGRADIENT_HPP
#define MYTGRADIENT_HPP

#include "MyTGL2/Util/MytFunctions.hpp"

namespace Myt{
	template <class T,class T_float>
	class IGradient{
	public:
		virtual T Get(T_float x)=0;
	};
	template <class T,class T_float,class T_Gradient>
	class IGradientImpl:public IGradient<T,T_float>,public T_Gradient{
	public:
		virtual T Get(T_float x){
			return static_cast<T_Gradient*>(this)->Get(x);
		}
	};

	template <class T,class T_float>
	class ConstantGradient{
	public:
		T Value;
	public:
		inline T Get(T_float x){
			return Value;
		}
	};

	template <class T,class T_float>
	class SimpleLinearGradient{
	public:
		T Value[2];
	public:
		inline T Get(T_float x){
			return LinearFunctions<T,T_float>::Lerp(Value[0],Value[1],x);
		}
	};
}

#endif