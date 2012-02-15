#include "MyTGL2/DataType/MytVector.hpp"
#include "MyTGL2/DataType/MytBitmap.hpp"
#include "MyTGL2/Util/MytBmpSaver.hpp"
#include "MyTGL2/Util/MytProfiler.hpp"
#include "MyTGL2/Util/MytFFT.hpp"
#include "MyTGL2/Operators/MytBitmapFlat.hpp"
#include "MyTGL2/Operators/MytBitmapGradient.hpp"
#include "MyTGL2/Rnd/MytMT19937.hpp"
#include <stdio.h>

int main(){
	//test only
	Myt::Bitmap8UC4 bm(8,8);
	Myt::MT19937 rnd;
	Myt::Timing t;
	//===Gradient2 test
	t.Start();
	Myt::Operators::Gradient2Property<Myt::Vector8UC4,float> props={255,0,0,255,0,255,0,255,0,0,255,255,255,255,0,255};
	Myt::Operators::Gradient2(&bm,&props);
	t.Stop();
	Myt::SaveBmp(bm,"out0.bmp");
	bm.Create(9,9);
	//
	printf("Time=%0.3fms\n",t.GetMs());
	//
	t.Clear();
	t.Start();
	//===FBM test
	rnd.Init(125);
	Myt::Complex32F *buf=Myt::Malloc<Myt::Complex32F>(bm.Width()*(bm.Height()/2+1)),*lp=buf;
	for(int j=0;j<bm.Height()/2+1;j++){
		for(int i=0;i<bm.Width();i++){
			int ii=Myt::OrderFunctions<int>::Min(i,bm.Width()-i);
			float d=Myt::OrderFunctions<float>::Max(sqrt(float(ii*ii+j*j)),1.0f);
			*lp=(Myt::Random<Myt::Complex32F>::Rnd(rnd)
				-Myt::Complex32F::Make(0.5f,0.5f))/d/sqrt(d); //*pow(d,-1.5f);
			lp++;
		}
	}
	Myt::FFTShuffleProvider sp[2];
	sp[0].Create(bm.Width());
	sp[1].Create(bm.Height()/2);
	Myt::FFT<float,float>::Calc2DReal(buf,buf,sp[0],sp[1],true,2);
	float m,n;
	Myt::OrderFunctions<float>::FindMinMax((float*)buf,m,n,bm.Width()*bm.Height());
	for(int i=0;i<bm.Width()*bm.Height();i++){
		bm[i]=(unsigned char)((((float*)buf)[i]-m)/(n-m)*255.0f);
	}
	Myt::Free(buf);
	//===
	t.Stop();
	Myt::SaveBmp(bm,"out.bmp");
	printf("Time=%0.3fms\n",t.GetMs());
	return 0;
}