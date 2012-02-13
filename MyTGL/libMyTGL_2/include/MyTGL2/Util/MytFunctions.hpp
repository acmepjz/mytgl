#ifndef MYTFUNCTIONS_HPP
#define MYTFUNCTIONS_HPP

#include <math.h>

namespace Myt{
	template<class T>
	struct Functions{
		static inline T Zero(){
			return 0;
		}
		static inline T One(){
			return 1;
		}
		static inline T Clamp(const T& Value,const T& Min,const T& Max){
			return Value<Min?Min:(Value>Max?Max:Value);
		}
	};
}

#endif