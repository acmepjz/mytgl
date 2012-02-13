#ifndef MYTBITMAPFLAT_HPP
#define MYTBITMAPFLAT_HPP

namespace Myt{
	namespace Operators{
		namespace Bitmap{
			//TEST ONLY
			template<class T_Color>
			struct FlatProperty{
				T_Color m_Color;
				inline T_Color& Color(){
					return m_Color;
				}
			};
			//TEST ONLY
			template<class T_Bitmap,class T_Property>
			inline void Flat(T_Bitmap* out,void* in_,T_Property* props){
				//T_Bitmap* in=(T_Bitmap*)in_;
				unsigned int m=1<<(out->WidthShift()+out->HeightShift());
				for(unsigned int i=0;i<m;i++){
					(*out)[i]=props->Color();
				}
			}
		}
	}
}

#endif