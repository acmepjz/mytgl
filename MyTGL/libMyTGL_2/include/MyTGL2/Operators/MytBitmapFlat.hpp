#ifndef MYTBITMAPFLAT_HPP
#define MYTBITMAPFLAT_HPP

#include "MyTGL2/DataType/MytBitmap.hpp"

namespace Myt{
	namespace Operators{
		template<class T_Color>
		struct FlatProperty{
			T_Color Color;
			inline T_Color GetColor(){
				return Color;
			}
		};

		template<class T,class T_Property>
		inline void Flat(Bitmap<T>* out,T_Property* props){
			unsigned int m=1UL<<(out->WidthShift()+out->HeightShift());
			for(unsigned int i=0;i<m;i++){
				(*out)[i]=props->GetColor();
			}
		}
	}
}

#endif