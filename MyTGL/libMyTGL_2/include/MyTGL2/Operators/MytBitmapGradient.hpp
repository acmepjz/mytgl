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

		template<class T,class T_float,class T_Gradient=IGradient<T,T_float> >
		struct Gradient2Property2{
			typedef T_float FloatType;

			int Direction;
			T_Gradient* Color[2];

			inline int GetDirection(){
				return Direction;
			}
			inline T GetColor0(T_float f){
				return Color[0]->Get(f);
			}
			inline T GetColor1(T_float f){
				return Color[1]->Get(f);
			}
			inline T GetColor(const T& obj1,const T& obj2,T_float f){
				return LinearFunctions<T,T_float>::Lerp(obj1,obj2,f);
			}
		};

		template<class T,class T_float,class T_Gradient=IGradient<T,T_float>,class T_Curve=IGradient<T_float,T_float> >
		struct Gradient2Property3{
			typedef T_float FloatType;

			int Direction;
			T_Gradient* Color[2];
			T_Curve* Curve;

			inline int GetDirection(){
				return Direction;
			}
			inline T GetColor0(T_float f){
				return Color[0]->Get(f);
			}
			inline T GetColor1(T_float f){
				return Color[1]->Get(f);
			}
			inline T GetColor(const T& obj1,const T& obj2,T_float f){
				return LinearFunctions<T,T_float>::Lerp(obj1,obj2,Curve->Get(f));
			}
		};

		template<class T,class T_Property>
		inline void Gradient2(Bitmap<T>* out,T_Property* props){
			unsigned int w=1UL<<(out->WidthShift()),h=1UL<<(out->HeightShift());
			T* lp=out->Pointer();
			T_Property::FloatType fw=T_Property::FloatType(1)/T_Property::FloatType(w),
				fh=T_Property::FloatType(1)/T_Property::FloatType(h);
			if(props->GetDirection()==0){
				T_Property::FloatType x=T_Property::FloatType(0);
				for(unsigned int i=0;i<w;i++){
					T clr0=props->GetColor0(x),clr1=props->GetColor1(x);
					T* lp1=lp+i;
					T_Property::FloatType y=T_Property::FloatType(0);
					for(unsigned int j=0;j<h;j++){
						*lp1=props->GetColor(clr0,clr1,y);
						lp1+=w;
						y+=fh;
					}
					x+=fw;
				}
			}else{
				T_Property::FloatType y=T_Property::FloatType(0);
				for(unsigned int j=0;j<h;j++){
					T clr0=props->GetColor0(y),clr1=props->GetColor1(y);
					T_Property::FloatType x=T_Property::FloatType(0);
					for(unsigned int i=0;i<w;i++){
						*(lp++)=props->GetColor(clr0,clr1,x);
						x+=fw;
					}
					y+=fh;
				}
			}
		}
	}
}

#endif
