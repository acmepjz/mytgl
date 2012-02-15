#ifndef MYTBITMAPGRADIENT_HPP
#define MYTBITMAPGRADIENT_HPP

#include "MyTGL2/Util/MytFunctions.hpp"
#include "MyTGL2/DataType/MytBitmap.hpp"
#include "MyTGL2/DataType/MytGradient.hpp"

namespace Myt{
	namespace Operators{
		template<class T,class T_float>
		struct Gradient2Property{
			typedef T_float FloatType;

			T Color[4];
			inline int GetDirection(){
				return 1;
			}
			inline T GetColor0(T_float f){
				return LinearFunctions<T,T_float>::Lerp(Color[0],Color[2],f);
			}
			inline T GetColor1(T_float f){
				return LinearFunctions<T,T_float>::Lerp(Color[1],Color[3],f);
			}
			inline T GetColor(const T& obj1,const T& obj2,T_float f){
				return LinearFunctions<T,T_float>::Lerp(obj1,obj2,f);
			}
		};

		template<class T,class T_Property>
		inline void Gradient2(Bitmap<T>* out,T_Property* props){
			unsigned int w=1UL<<(out->WidthShift()),h=1UL<<(out->HeightShift());
			T* lp=out->Pointer();
			T_Property::FloatType fw=1.0f/T_Property::FloatType(w),fh=1.0f/T_Property::FloatType(h);
			if(props->GetDirection()==0){
				for(unsigned int i=0;i<w;i++){
					T_Property::FloatType x=T_Property::FloatType(i)*fw;
					T clr0=props->GetColor0(x),clr1=props->GetColor1(x);
					T* lp1=lp+i;
					for(unsigned int j=0;j<h;j++){
						T_Property::FloatType y=T_Property::FloatType(j)*fh;
						*lp1=props->GetColor(clr0,clr1,y);
						lp1+=w;
					}
				}
			}else{
				for(unsigned int j=0;j<h;j++){
					T_Property::FloatType y=T_Property::FloatType(j)*fh;
					T clr0=props->GetColor0(y),clr1=props->GetColor1(y);
					for(unsigned int i=0;i<w;i++){
						T_Property::FloatType x=T_Property::FloatType(i)*fw;
						*(lp++)=props->GetColor(clr0,clr1,x);
					}
				}
			}
		}
	}
}

#endif
