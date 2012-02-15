#ifndef MYTBITMAP_HPP
#define MYTBITMAP_HPP

#include "MyTGL2/Util/MytMemoryManagement.hpp"
#include "MyTGL2/DataType/MytVector.hpp"

namespace Myt{
	template<class T_DataType>
	class Bitmap{
	private:
		unsigned char ws,hs;
		T_DataType *lp;
	public:
		typedef T_DataType DataType;

		inline Bitmap(){
			ws=0;
			hs=0;
			lp=NULL;
		}

		inline Bitmap(unsigned char WidthShift_,unsigned char HeightShift_){
			ws=WidthShift_;
			hs=HeightShift_;
			lp=AlignedMalloc<T_DataType,16>(1<<(ws+hs));
		}

		inline Bitmap(unsigned char WidthShift_,unsigned char HeightShift_,const void* lpSrc){
			ws=WidthShift_;
			hs=HeightShift_;
			lp=AlignedMalloc<T_DataType,16>(1<<(ws+hs));
			memcpy(lp,lpSrc,sizeof(T_DataType)<<(ws+hs));
		}

		inline Bitmap(const Bitmap<T_DataType>& obj){
			ws=obj.ws;
			hs=obj.hs;
			lp=AlignedMalloc<T_DataType,16>(1<<(ws+hs));
			memcpy(lp,obj.lp,sizeof(T_DataType)<<(ws+hs));
		}

		inline Bitmap<T_DataType>& operator=(const Bitmap<T_DataType>& obj){
			if(ws+hs!=obj.ws+obj.hs || lp==NULL){
				if(lp) AlignedFree(lp);
				lp=AlignedMalloc<T_DataType,16>(1<<(obj.ws+obj.hs));
			}
			ws=obj.ws;
			hs=obj.hs;
			memcpy(lp,obj.lp,sizeof(T_DataType)<<(ws+hs));
		}

		inline void Destroy(){
			ws=0;
			hs=0;
			if(lp) AlignedFree(lp);
			lp=NULL;
		}

		inline void Create(unsigned char WidthShift_,unsigned char HeightShift_){
			if(ws+hs!=WidthShift_+HeightShift_ || lp==NULL){
				if(lp) AlignedFree(lp);
				lp=AlignedMalloc<T_DataType,16>(1<<(WidthShift_+HeightShift_));
			}
			ws=WidthShift_;
			hs=HeightShift_;
		}

		~Bitmap(){
			AlignedFree(lp);
		}

		//IBitmap must implement
		inline unsigned char WidthShift() const{
			return ws;
		}
		//IBitmap must implement
		inline unsigned char HeightShift() const{
			return hs;
		}
		//IBitmap must implement (one of them)
		inline T_DataType* Pointer(){
			return lp;
		}
		inline const T_DataType* Pointer() const{
			return lp;
		}
		//IBitmap must implement (one of them)
		inline const T_DataType& operator[](int idx) const{
			return lp[idx];
		}
		inline T_DataType& operator[](int idx){
			return lp[idx];
		}
		//IBitmap must implement (one of them)
		inline const T_DataType& operator()(int x,int y) const{
			return lp[(y<<ws)+x];
		}
		inline T_DataType& operator()(int x,int y){
			return lp[(y<<ws)+x];
		}

		inline int Width() const{
			return 1<<ws;
		}
		inline int Height() const{
			return 1<<hs;
		}
		inline void Clear(){
			memset(lp,0,sizeof(T_DataType)<<(ws+hs));
		}
	};

	typedef Bitmap<Vector8UC1> Bitmap8UC1;
	typedef Bitmap<Vector8UC2> Bitmap8UC2;
	typedef Bitmap<Vector8UC3> Bitmap8UC3;
	typedef Bitmap<Vector8UC4> Bitmap8UC4;

	typedef Bitmap<Vector32FC1> Bitmap32FC1;
	typedef Bitmap<Vector32FC2> Bitmap32FC2;
	typedef Bitmap<Vector32FC3> Bitmap32FC3;
	typedef Bitmap<Vector32FC4> Bitmap32FC4;
}

#endif
