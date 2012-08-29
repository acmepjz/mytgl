#include "MyTGL2/DataType/MytVector.hpp"
#include "MyTGL2/DataType/MytBitmap.hpp"
#include "MyTGL2/DataType/MytGradient.hpp"
#include "MyTGL2/Util/MytBmpSaver.hpp"
#include "MyTGL2/Util/MytProfiler.hpp"
#include "MyTGL2/Util/MytFFT.hpp"
#include "MyTGL2/Util/MytPerlinNoise.hpp"
#include "MyTGL2/Operators/MytBitmapFlat.hpp"
#include "MyTGL2/Operators/MytBitmapGradient.hpp"
#include "MyTGL2/Rnd/MytMT19937.hpp"
#include "MyTGL2/Rnd/MytSimpleNoise.hpp"
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
	printf("Time=%0.3fms\n",t.GetMs());
	//===another test
	t.Clear();
	t.Start();
	{
		rnd.Init(196);
		Myt::IGradientImpl<Myt::Vector8UC4,float,Myt::SequentialLinearGradient<Myt::Vector8UC4,float> > grad1,grad2;
		Myt::IGradientImpl<float,float,Myt::SequentialLinearGradient<float,float> > grad3;
		for(int i=0;i<10;i++){
			Myt::LinearGradient<Myt::Vector8UC4,float>::PointType p={Myt::Random<float>::Rnd(rnd),Myt::Random<Myt::Vector8UC4>::Rnd(rnd)},
				p2={Myt::Random<float>::Rnd(rnd),Myt::Random<Myt::Vector8UC4>::Rnd(rnd)};
			p.y[3]=255;
			p2.y[3]=255;
			grad1.Insert(p);
			grad2.Insert(p2);
			Myt::LinearGradient<float,float>::PointType p3={Myt::Random<float>::Rnd(rnd)*0.75f+0.125f,Myt::Random<float>::Rnd(rnd)};
			grad3.Insert(p3);
		}
		Myt::LinearGradient<float,float>::PointType p3={0,0};
		grad3.Insert(p3);
		p3.x=1;p3.y=1;
		grad3.Insert(p3);
		Myt::Operators::Gradient2Property3<Myt::Vector8UC4,float> props={0,&grad1,&grad2,&grad3};
		Myt::Operators::Gradient2(&bm,&props);
	}
	t.Stop();
	Myt::SaveBmp(bm,"out0b.bmp");
	printf("Time=%0.3fms\n",t.GetMs());
	//===perlin test
	bm.Create(9,9);
	Myt::SimpleNoise noise;
	t.Clear();
	t.Start();
	for(int j=0;j<bm.Height();j++){
		for(int i=0;i<bm.Width();i++){
			float f=0.0f;
			float f1=1.0f/64.0f,f2=1.0f;
			for(int k=0;k<6;k++){
				f+=Myt::PerlinNoise<Myt::SimpleNoise>::Noise2(float(i)*f1,float(j)*f1,345+k,noise)*f2;
				f1*=2.0f;
				f2*=0.5f;
			}
			bm(i,j)=(unsigned char)Myt::OrderFunctions<float>::Clamp(
				f*128.0f+128.0f,
				0.0f,255.0f);
		}
	}
	t.Stop();
	Myt::SaveBmp(bm,"out0c.bmp");
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
			float d=Myt::OrderFunctions<float>::Max(sqrt(float(ii*ii+j*j)),5.0f);
			*lp=(Myt::Random<Myt::Complex32F>::Rnd(rnd)
				-Myt::Complex32F::Make(0.5f,0.5f))*pow(d,-2.00f);
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