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
	};
}

#endif