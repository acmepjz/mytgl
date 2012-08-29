#ifndef MYTBMPSAVER_HPP
#define MYTBMPSAVER_HPP

#include "MyTGL2/DataType/MytVector.hpp"
#include "MyTGL2/DataType/MytBitmap.hpp"
#include <stdio.h>

namespace Myt{
	inline bool SaveBmp(const Bitmap8UC4& bm,FILE* f){
		int Size=4<<(bm.WidthShift()+bm.HeightShift());
		int Size1=Size+54;
		int Width=bm.Width();
		int Height=-bm.Height();
		unsigned char b[54]={0};
		/*bfType     */ b[0]='B';    b[1]='M';
		/*bfSize     */ b[2]=Size1;  b[3]=Size1>>8;  b[4]=Size1>>16;  b[5]=Size1>>24;
		/*bfOffBits  */ b[10]=54;
		/*biSize     */ b[14]=40;
		/*biWidth    */ b[18]=Width; b[19]=Width>>8; b[20]=Width>>16; b[21]=Width>>24;
		/*biHeight   */ b[22]=Height;b[23]=Height>>8;b[24]=Height>>16;b[25]=Height>>24;
		/*biPlanes   */ b[26]=1;
		/*biBitCount */ b[28]=32;
		/*biSizeImage*/ b[34]=Size;  b[35]=Size>>8;  b[36]=Size>>16;  b[37]=Size>>24;
		fwrite(b,1,54,f);
		fwrite(bm.Pointer(),1,Size,f);
		return true;
	}
	template <class T_Bmp>
	inline bool SaveBmp(const T_Bmp& bm,const char* fn){
		FILE *f=fopen(fn,"wb");
		bool b=false;
		if(f){
			b=SaveBmp(bm,f);
			fclose(f);
		}
		return b;
	}
}

#endif