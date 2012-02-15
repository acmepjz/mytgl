#ifndef MYTFFT_HPP
#define MYTFFT_HPP

#include "MyTGL2/Util/MytMemoryManagement.hpp"
#include "MyTGL2/Util/MytFunctions.hpp"
#include "MyTGL2/DataType/MytVector.hpp"
#include "MyTGL2/DataType/MytComplex.hpp"
#include <math.h>

namespace Myt{
	struct FFTShuffleProvider{
	public:
		unsigned int N;
		unsigned int Shift;
		unsigned int* Target;
	public:
		inline FFTShuffleProvider(){
			N=0;
			Shift=0;
			Target=NULL;
		}
		//N should be power-of-two.
		inline void Create(unsigned int N_){
			if(N_==N) return;
			else{
				Myt::Free(Target);
				Target=Myt::Malloc<unsigned int>(N_);
				N=N_;
				Shift=((N_&0xAAAAAAA)?1:0)|((N_&0xCCCCCCCC)?2:0)|((N_&0xF0F0F0F0)?4:0)|((N_&0xFF00FF00)?8:0)|((N_&0xFFFF0000)?16:0);
			}

			unsigned int t=0;

			for(unsigned int p=0;p<N_;p++){
				Target[p]=t;

				unsigned int Mask=N_;
				while(t & (Mask>>=1)) t&=~Mask;

				t|=Mask;
			}
		}
		~FFTShuffleProvider(){
			Myt::Free(Target);
		}
	};
	//Only supports power-of-two data size.
	//transfer complex data to complex data.
	template<class T,class T_float>
	class FFT{
	public:
		static void Calc(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,bool Inverse,int Normalize){
			if(Src==Dest) Shuffle(Dest,N);
			else Shuffle(Src,Dest,N);
			CalcInternalAndPostProcess(Dest,N,Inverse,Normalize);
		}
		static void Calc(const Complex<T>* Src,Complex<T>* Dest,const FFTShuffleProvider& obj,bool Inverse,int Normalize){
			const unsigned int N=obj.N;
			if(Src==Dest) Shuffle(Dest,N,obj);
			else Shuffle(Src,Dest,N,obj);
			CalcInternalAndPostProcess(Dest,N,Inverse,Normalize);
		}
	private:
		static inline void CalcInternalAndPostProcess(Complex<T>* Dest,unsigned int N,bool Inverse,int Normalize){
			CalcInternal(Dest,N,Inverse);
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}else if(Inverse && Normalize==0){
				T_float f=T_float(1.0)/T_float(N);
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}
		}
		static void Shuffle(const Complex<T>* Src,Complex<T>* Dest,unsigned int N){
			unsigned int Target=0;

			for(unsigned int p=0;p<N;p++){
				Dest[Target]=Src[p];

				unsigned int Mask=N;
				while(Target & (Mask>>=1)) Target&=~Mask;

				Target|=Mask;
			}
		}
		static void Shuffle(Complex<T>* Dest,unsigned int N){
			unsigned int Target = 0;

			for(unsigned int p=0;p<N;p++){
				if(Target>p){
					const Complex<T> tmp=Dest[Target];
					Dest[Target]=Dest[p];
					Dest[p]=tmp;
				}

				unsigned int Mask=N;
				while(Target & (Mask>>=1)) Target&=~Mask;

				Target|=Mask;
			}
		}
		static void Shuffle(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,const FFTShuffleProvider& obj){
			for(unsigned int p=0;p<N;p++){
				unsigned int Target=obj.Target[p];
				Dest[Target]=Src[p];
			}
		}
		static void Shuffle(Complex<T>* Dest,unsigned int N,const FFTShuffleProvider& obj){
			for(unsigned int p=0;p<N;p++){
				unsigned int Target=obj.Target[p];
				if(Target>p){
					const Complex<T> tmp=Dest[Target];
					Dest[Target]=Dest[p];
					Dest[p]=tmp;
				}
			}
		}
		static void CalcInternal(Complex<T>* Data,unsigned int N,bool Inverse){
			const T_float pi=Inverse?T_float(3.14159265358979323846):T_float(-3.14159265358979323846);

			for(unsigned int Step=1;Step<N;Step<<=1){
				if((Step<<1)<N){
					//butterfly4
					const unsigned int Jump=Step<<2;

					const T_float delta=pi/T_float(Step<<1);
					const Complex<T_float> Multiplier={cos(delta),sin(delta)};
					const Complex<T_float> Multiplier2={cos(delta*T_float(2.0)),sin(delta*T_float(2.0))};
					const Complex<T_float> Multiplier3={cos(delta*T_float(3.0)),sin(delta*T_float(3.0))};
					Complex<T_float> Factor={1,0};
					Complex<T_float> Factor2={1,0};
					Complex<T_float> Factor3={1,0};

					for(unsigned int i=0;i<Step;i++){
						for(unsigned int j=i;j<N;j+=Jump){
							const unsigned int k=j+Step,k2=j+Step*2,k3=j+Step*3;
							Complex<T> tmp[6];

							//note that the data is radix-2 shuffled
							tmp[0]=Data[k2]*Factor;
							tmp[1]=Data[k]*Factor2;
							tmp[2]=Data[k3]*Factor3;

							tmp[5]=Data[j]-tmp[1];
							Data[j]+=tmp[1];
							tmp[3]=tmp[0]+tmp[2];
							tmp[4]=tmp[0]-tmp[2];
							Data[k2]=Data[j]-tmp[3];

							Data[j]+=tmp[3];

							if(Inverse){
								Data[k].Re=tmp[5].Re-tmp[4].Im;
								Data[k].Im=tmp[5].Im+tmp[4].Re;
								Data[k3].Re=tmp[5].Re+tmp[4].Im;
								Data[k3].Im=tmp[5].Im-tmp[4].Re;
							}else{
								Data[k].Re=tmp[5].Re+tmp[4].Im;
								Data[k].Im=tmp[5].Im-tmp[4].Re;
								Data[k3].Re=tmp[5].Re-tmp[4].Im;
								Data[k3].Im=tmp[5].Im+tmp[4].Re;
							}
						}
						Factor*=Multiplier;
						Factor2*=Multiplier2;
						Factor3*=Multiplier3;
					}

					Step<<=1;
				}else{
					//butterfly2
					const unsigned int Jump=Step<<1;

					const T_float delta=pi/T_float(Step);
					const Complex<T_float> Multiplier={cos(delta),sin(delta)};
					Complex<T_float> Factor={1,0};

					for(unsigned int i=0;i<Step;i++){
						for(unsigned int j=i;j<N;j+=Jump){
							const unsigned int k=j+Step;
							const Complex<T> Product=Data[k]*Factor;
							Data[k]=Data[j]-Product;
							Data[j]+=Product;
						}
						Factor*=Multiplier;
					}
				}
			}
		}
	public:
		//N must be even, power-of-two
		//output size is N/2 complex number, Dest[N/2].Re is put in Dest[0].Im
		static void CalcReal(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,bool Inverse,int Normalize){
			if(Inverse){
				CalcRealInversePreProcess(Src,Dest,N>>1);
				Shuffle(Dest,N>>1);
				CalcRealInverseInternalAndPostProcess(Dest,N>>1,Normalize);
			}else{
				if(Src==Dest) Shuffle(Dest,N>>1);
				else Shuffle(Src,Dest,N>>1);
				CalcRealInternalAndPostProcess(Dest,N>>1,Normalize);
			}
		}
		//real data size should be 2*obj.N
		//output size is obj.N complex number, Dest[obj.N].Re is put in Dest[0].Im
		static void CalcReal(const Complex<T>* Src,Complex<T>* Dest,const FFTShuffleProvider& obj,bool Inverse,int Normalize){
			if(Inverse){
				CalcRealInversePreProcess(Src,Dest,obj.N);
				Shuffle(Dest,obj.N,obj);
				CalcRealInverseInternalAndPostProcess(Dest,obj.N,Normalize);
			}else{
				if(Src==Dest) Shuffle(Dest,obj.N,obj);
				else Shuffle(Src,Dest,obj.N,obj);
				CalcRealInternalAndPostProcess(Dest,obj.N,Normalize);
			}
		}
	private:
		static void CalcRealInternalAndPostProcess(Complex<T>* Dest,unsigned int N,int Normalize){
			CalcInternal(Dest,N,false);
			//post process
			Dest[0]=Complex<T>::Make(
				Dest[0].Re+Dest[0].Im, //Dest[0]
				Dest[0].Re-Dest[0].Im  //Dest[N]
				);

			const T_float pi=T_float(-3.14159265358979323846);
			const T_float delta=pi/T_float(N);
			const Complex<T_float> Multiplier={cos(delta),sin(delta)};
			Complex<T_float> Factor={0,-1};

			for(unsigned int k=1;k<=(N>>1);k++){
				Factor*=Multiplier;

				Complex<T> tmp[2];
				tmp[0]=!Dest[N-k];

				tmp[1]=(Dest[k]-tmp[0])*Factor;
				tmp[0]+=Dest[k];

				Dest[k]=(tmp[0]+tmp[1])*T_float(0.5);
				Dest[N-k]=Complex<T>::Make(
					(tmp[0].Re-tmp[1].Re)*T_float(0.5),
					(tmp[1].Im-tmp[0].Im)*T_float(0.5));
			}
			//over
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N<<1));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}
		}
		static void CalcRealInversePreProcess(const Complex<T>* Src,Complex<T>* Dest,unsigned int N){
			Dest[0]=Complex<T>::Make(
				Src[0].Re+Src[0].Im,
				Src[0].Re-Src[0].Im);

			const T_float pi=T_float(3.14159265358979323846);
			const T_float delta=pi/T_float(N);
			const Complex<T_float> Multiplier={cos(delta),sin(delta)};
			Complex<T_float> Factor={0,1};

			for(unsigned int k=1;k<=(N>>1);k++){
				Factor*=Multiplier;

				Complex<T> tmp[2];
				tmp[0]=!Src[N-k];

				tmp[1]=(Src[k]-tmp[0])*Factor;
				tmp[0]+=Src[k];

				Dest[k]=tmp[0]+tmp[1];
				Dest[N-k]=Complex<T>::Make(
					tmp[0].Re-tmp[1].Re,
					tmp[1].Im-tmp[0].Im);
			}
		}
		static inline void CalcRealInverseInternalAndPostProcess(Complex<T>* Dest,unsigned int N,int Normalize){
			CalcInternal(Dest,N,true);
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N<<1));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}else if(Normalize==0){
				T_float f=T_float(1.0)/T_float(N<<1);
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}
		}
	public:
		static void CalcND(const Complex<T>* Src,Complex<T>* Dest,unsigned int Dimension,const FFTShuffleProvider* obj,bool Inverse,int Normalize){
			//calc size
			unsigned int N=1UL;
			for(unsigned int d=0;d<Dimension;d++) N<<=obj[d].Shift;
			//calc d=0
			unsigned int Step=obj[0].N,BlockSizeShift=obj[0].Shift;
			if(Src==Dest){
				for(unsigned int i=0;i<N;i+=Step){
					Shuffle(Dest+i,Step,obj[0]);
					CalcInternal(Dest+i,Step,Inverse);
				}
			}else{
				for(unsigned int i=0;i<N;i+=Step){
					Shuffle(Src+i,Dest+i,Step,obj[0]);
					CalcInternal(Dest+i,Step,Inverse);
				}
			}
			//calc others
			for(unsigned int d=1;d<Dimension;d++){
				Step<<=obj[d].Shift;
				for(unsigned int i=0;i<N;i+=Step){
					ShuffleEx(Dest+i,obj[d].N,obj[d],BlockSizeShift);
					CalcInternalEx(Dest+i,obj[d].N,Inverse,BlockSizeShift);
				}
				BlockSizeShift+=obj[d].Shift;
			}
			//over
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}else if(Inverse && Normalize==0){
				T_float f=T_float(1.0)/T_float(N);
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}
		}
		static void Calc2D(const Complex<T>* Src,Complex<T>* Dest,const FFTShuffleProvider& obj0,const FFTShuffleProvider& obj1,bool Inverse,int Normalize){
			//calc size
			unsigned int N=1UL<<(obj0.Shift+obj1.Shift);
			//calc d=0
			unsigned int Step=obj0.N,BlockSizeShift=obj0.Shift;
			if(Src==Dest){
				for(unsigned int i=0;i<N;i+=Step){
					Shuffle(Dest+i,Step,obj0);
					CalcInternal(Dest+i,Step,Inverse);
				}
			}else{
				for(unsigned int i=0;i<N;i+=Step){
					Shuffle(Src+i,Dest+i,Step,obj0);
					CalcInternal(Dest+i,Step,Inverse);
				}
			}
			//calc d=1
			ShuffleEx(Dest,obj1.N,obj1,BlockSizeShift);
			CalcInternalEx(Dest,obj1.N,Inverse,BlockSizeShift);
			//over
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}else if(Inverse && Normalize==0){
				T_float f=T_float(1.0)/T_float(N);
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}
		}
	private:
		static void ShuffleEx(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,const FFTShuffleProvider& obj,unsigned int BlockSizeShift){
			for(unsigned int p=0;p<N;p++){
				unsigned int Target=obj.Target[p];
				for(unsigned int idx=0;idx<(1UL)<<BlockSizeShift;idx++){
					Dest[(Target<<BlockSizeShift)+idx]=Src[(p<<BlockSizeShift)+idx];
				}
			}
		}
		static void ShuffleEx(Complex<T>* Dest,unsigned int N,const FFTShuffleProvider& obj,unsigned int BlockSizeShift){
			for(unsigned int p=0;p<N;p++){
				unsigned int Target=obj.Target[p];
				if(Target>p){
					for(unsigned int idx=0;idx<(1UL)<<BlockSizeShift;idx++){
						const Complex<T> tmp=Dest[(Target<<BlockSizeShift)+idx];
						Dest[(Target<<BlockSizeShift)+idx]=Dest[(p<<BlockSizeShift)+idx];
						Dest[(p<<BlockSizeShift)+idx]=tmp;
					}
				}
			}
		}
		static void CalcInternalEx(Complex<T>* Data,unsigned int N,bool Inverse,unsigned int BlockSizeShift){
			const T_float pi=Inverse?T_float(3.14159265358979323846):T_float(-3.14159265358979323846);

			for(unsigned int Step=1;Step<N;Step<<=1){
				if((Step<<1)<N){
					//butterfly4
					const unsigned int Jump=Step<<2;

					const T_float delta=pi/T_float(Step<<1);
					const Complex<T_float> Multiplier={cos(delta),sin(delta)};
					const Complex<T_float> Multiplier2={cos(delta*T_float(2.0)),sin(delta*T_float(2.0))};
					const Complex<T_float> Multiplier3={cos(delta*T_float(3.0)),sin(delta*T_float(3.0))};
					Complex<T_float> Factor={1,0};
					Complex<T_float> Factor2={1,0};
					Complex<T_float> Factor3={1,0};

					for(unsigned int i=0;i<Step;i++){
						for(unsigned int j=i;j<N;j+=Jump){
							for(unsigned int idx=0;idx<(1UL)<<BlockSizeShift;idx++){
								const unsigned int k0=(j<<BlockSizeShift)+idx,
									k=((j+Step)<<BlockSizeShift)+idx,
									k2=((j+Step*2)<<BlockSizeShift)+idx,
									k3=((j+Step*3)<<BlockSizeShift)+idx;
								Complex<T> tmp[6];

								//note that the data is radix-2 shuffled
								tmp[0]=Data[k2]*Factor;
								tmp[1]=Data[k]*Factor2;
								tmp[2]=Data[k3]*Factor3;

								tmp[5]=Data[k0]-tmp[1];
								Data[k0]+=tmp[1];
								tmp[3]=tmp[0]+tmp[2];
								tmp[4]=tmp[0]-tmp[2];
								Data[k2]=Data[k0]-tmp[3];

								Data[k0]+=tmp[3];

								if(Inverse){
									Data[k].Re=tmp[5].Re-tmp[4].Im;
									Data[k].Im=tmp[5].Im+tmp[4].Re;
									Data[k3].Re=tmp[5].Re+tmp[4].Im;
									Data[k3].Im=tmp[5].Im-tmp[4].Re;
								}else{
									Data[k].Re=tmp[5].Re+tmp[4].Im;
									Data[k].Im=tmp[5].Im-tmp[4].Re;
									Data[k3].Re=tmp[5].Re-tmp[4].Im;
									Data[k3].Im=tmp[5].Im+tmp[4].Re;
								}
							}
						}
						Factor*=Multiplier;
						Factor2*=Multiplier2;
						Factor3*=Multiplier3;
					}

					Step<<=1;
				}else{
					//butterfly2
					const unsigned int Jump=Step<<1;

					const T_float delta=pi/T_float(Step);
					const Complex<T_float> Multiplier={cos(delta),sin(delta)};
					Complex<T_float> Factor={1,0};

					for(unsigned int i=0;i<Step;i++){
						for(unsigned int j=i;j<N;j+=Jump){
							for(unsigned int idx=0;idx<(1UL)<<BlockSizeShift;idx++){
								const unsigned int k0=(j<<BlockSizeShift)+idx,
									k=((j+Step)<<BlockSizeShift)+idx;
								const Complex<T> Product=Data[k]*Factor;
								Data[k]=Data[k0]-Product;
								Data[k0]+=Product;
							}
						}
						Factor*=Multiplier;
					}
				}
			}
		}
	public:
		//obj[Dimension-1].N must be even
		//real data size: obj[0].N * ... * (obj[Dimension-1].N*2)
		//complex data size: obj[0].N * ... * (obj[Dimension-1].N+1)
		//if Dimension>1 && Inverse==true then the output real data size should be
		//as large as complex data, in order to hold temp data
		static void CalcNDReal(const Complex<T>* Src,Complex<T>* Dest,unsigned int Dimension,const FFTShuffleProvider* obj,bool Inverse,int Normalize){
			//calc size
			unsigned int S0=0;
			for(unsigned int d=0;d<Dimension-1;d++) S0+=obj[d].Shift;

			unsigned int N=((1UL<<S0)<<obj[Dimension-1].Shift)<<1; //real data size
			unsigned int N1=((1UL<<S0)<<obj[Dimension-1].Shift)+(1UL<<S0); //complex data size
			//calc d=Dimension-1 (forward)
			if(!Inverse){
				if(Dimension>1){
					if(Src==Dest) ShuffleNDRealInplace(Dest,Dest+(N>>1),obj[Dimension-1].N,S0);
					else ShuffleNDReal(Src,Dest,obj[Dimension-1].N,S0);
				}
				ShuffleEx(Dest,obj[Dimension-1].N,obj[Dimension-1],S0);
				CalcInternalEx(Dest,obj[Dimension-1].N,Inverse,S0);
				CalcRealPostProcessEx(Dest,obj[Dimension-1].N,S0);
			}
			if(Dimension>1){
				//calc d=0
				unsigned int Step=obj[0].N,BlockSizeShift=obj[0].Shift;
				if(Inverse && Src!=Dest){
					for(unsigned int i=0;i<N1;i+=Step){
						Shuffle(Src+i,Dest+i,Step,obj[0]);
						CalcInternal(Dest+i,Step,Inverse);
					}
				}else{
					for(unsigned int i=0;i<N1;i+=Step){
						Shuffle(Dest+i,Step,obj[0]);
						CalcInternal(Dest+i,Step,Inverse);
					}
				}
				//calc others
				for(unsigned int d=1;d<Dimension-1;d++){
					Step<<=obj[d].Shift;
					for(unsigned int i=0;i<N1;i+=Step){
						ShuffleEx(Dest+i,obj[d].N,obj[d],BlockSizeShift);
						CalcInternalEx(Dest+i,obj[d].N,Inverse,BlockSizeShift);
					}
					BlockSizeShift+=obj[d].Shift;
				}
			}
			//calc d=Dimension-1 (inverse)
			if(Inverse){
				CalcRealInversePreProcessEx(Dimension>1?Dest:Src,Dest,obj[Dimension-1].N,S0);
				ShuffleEx(Dest,obj[Dimension-1].N,obj[Dimension-1],S0);
				CalcInternalEx(Dest,obj[Dimension-1].N,Inverse,S0);
				if(Dimension>1) ShuffleNDRealInverseInplace(Dest,Dest+(N>>1),obj[Dimension-1].N,S0);
			}
			//over
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<(Inverse?(N>>1):N1);i++) Dest[i]*=f;
			}else if(Inverse && Normalize==0){
				T_float f=T_float(1.0)/T_float(N);
				for(unsigned int i=0;i<(N>>1);i++) Dest[i]*=f;
			}
		}
		//obj1.N must be even
		//real data size: obj0.N * (obj1.N*2)
		//complex data size: obj0.N * (obj1.N+1)
		//if Inverse==true then the output real data size should be
		//as large as complex data, in order to hold temp data
		static void Calc2DReal(const Complex<T>* Src,Complex<T>* Dest,const FFTShuffleProvider& obj0,const FFTShuffleProvider& obj1,bool Inverse,int Normalize){
			//calc size
			unsigned int S0=obj0.Shift;
			unsigned int N=((1UL<<S0)<<obj1.Shift)<<1; //real data size
			unsigned int N1=((1UL<<S0)<<obj1.Shift)+(1UL<<S0); //complex data size
			//calc d=1 (forward)
			if(!Inverse){
				if(Src==Dest) ShuffleNDRealInplace(Dest,Dest+(N>>1),obj1.N,S0);
				else ShuffleNDReal(Src,Dest,obj1.N,S0);

				ShuffleEx(Dest,obj1.N,obj1,S0);
				CalcInternalEx(Dest,obj1.N,Inverse,S0);
				CalcRealPostProcessEx(Dest,obj1.N,S0);
			}
			//calc d=0
			unsigned int Step=obj0.N;
			if(Inverse && Src!=Dest){
				for(unsigned int i=0;i<N1;i+=Step){
					Shuffle(Src+i,Dest+i,Step,obj0);
					CalcInternal(Dest+i,Step,Inverse);
				}
			}else{
				for(unsigned int i=0;i<N1;i+=Step){
					Shuffle(Dest+i,Step,obj0);
					CalcInternal(Dest+i,Step,Inverse);
				}
			}
			//calc d=1 (inverse)
			if(Inverse){
				CalcRealInversePreProcessEx(Dest,Dest,obj1.N,S0);
				ShuffleEx(Dest,obj1.N,obj1,S0);
				CalcInternalEx(Dest,obj1.N,Inverse,S0);
				
				ShuffleNDRealInverseInplace(Dest,Dest+(N>>1),obj1.N,S0);
			}
			//over
			if(Normalize==1){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<(Inverse?(N>>1):N1);i++) Dest[i]*=f;
			}else if(Inverse && Normalize==0){
				T_float f=T_float(1.0)/T_float(N);
				for(unsigned int i=0;i<(N>>1);i++) Dest[i]*=f;
			}
		}
	private:
		static void CalcRealPostProcessEx(Complex<T>* Dest,unsigned int N,unsigned int BlockSizeShift){
			for(unsigned int idx=0;idx<(1UL<<BlockSizeShift);idx++){
				Dest[(N<<BlockSizeShift)+idx]=Complex<T>::Make(Dest[idx].Re-Dest[idx].Im,Constants<T>::Zero());
				Dest[idx]=Complex<T>::Make(Dest[idx].Re+Dest[idx].Im,Constants<T>::Zero());
			}

			const T_float pi=T_float(-3.14159265358979323846);
			const T_float delta=pi/T_float(N);
			const Complex<T_float> Multiplier={cos(delta),sin(delta)};
			Complex<T_float> Factor={0,-1};

			for(unsigned int k=1;k<=(N>>1);k++){
				Factor*=Multiplier;

				for(unsigned int idx=0;idx<(1UL<<BlockSizeShift);idx++){
					Complex<T> tmp[2];
					tmp[0]=!Dest[((N-k)<<BlockSizeShift)+idx];

					tmp[1]=(Dest[(k<<BlockSizeShift)+idx]-tmp[0])*Factor;
					tmp[0]+=Dest[(k<<BlockSizeShift)+idx];

					Dest[(k<<BlockSizeShift)+idx]=(tmp[0]+tmp[1])*T_float(0.5);
					Dest[((N-k)<<BlockSizeShift)+idx]=Complex<T>::Make(
						(tmp[0].Re-tmp[1].Re)*T_float(0.5),
						(tmp[1].Im-tmp[0].Im)*T_float(0.5));
				}
			}
		}
		static void CalcRealInversePreProcessEx(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,unsigned int BlockSizeShift){
			for(unsigned int idx=0;idx<(1UL<<BlockSizeShift);idx++){
				Dest[idx]=Complex<T>::Make(
					Src[idx].Re+Src[(N<<BlockSizeShift)+idx].Re,
					Src[idx].Re-Src[(N<<BlockSizeShift)+idx].Re);
			}

			const T_float pi=T_float(3.14159265358979323846);
			const T_float delta=pi/T_float(N);
			const Complex<T_float> Multiplier={cos(delta),sin(delta)};
			Complex<T_float> Factor={0,1};

			for(unsigned int k=1;k<=(N>>1);k++){
				Factor*=Multiplier;

				for(unsigned int idx=0;idx<(1UL<<BlockSizeShift);idx++){
					Complex<T> tmp[2];
					tmp[0]=!Src[((N-k)<<BlockSizeShift)+idx];

					tmp[1]=(Src[(k<<BlockSizeShift)+idx]-tmp[0])*Factor;
					tmp[0]+=Src[(k<<BlockSizeShift)+idx];

					Dest[(k<<BlockSizeShift)+idx]=tmp[0]+tmp[1];
					Dest[((N-k)<<BlockSizeShift)+idx]=Complex<T>::Make(
						tmp[0].Re-tmp[1].Re,
						tmp[1].Im-tmp[0].Im);
				}
			}
		}
		static void ShuffleNDReal(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,unsigned int BlockSizeShift){
			for(unsigned int i=0;i<N;i++){
				const T* Src1=(const T*)(Src+(i<<BlockSizeShift));
				T* Dest1=(T*)(Dest+(i<<BlockSizeShift));
				for(unsigned int p=0;p<(1UL<<BlockSizeShift);p++){
					Dest1[p<<1]=Src1[p];
					Dest1[(p<<1)+1]=Src1[p+(1UL<<BlockSizeShift)];
				}
			}
		}
		static void ShuffleNDRealInplace(Complex<T>* Dest,Complex<T>* Temp,unsigned int N,unsigned int BlockSizeShift){
			for(unsigned int i=0;i<N;i++){
				T* Src1=(T*)Temp;
				T* Dest1=(T*)(Dest+(i<<BlockSizeShift));
				for(unsigned int p=0;p<(2UL<<BlockSizeShift);p++) Src1[p]=Dest1[p];
				for(unsigned int p=0;p<(1UL<<BlockSizeShift);p++){
					Dest1[p<<1]=Src1[p];
					Dest1[(p<<1)+1]=Src1[p+(1UL<<BlockSizeShift)];
				}
			}
		}
		static void ShuffleNDRealInverseInplace(Complex<T>* Dest,Complex<T>* Temp,unsigned int N,unsigned int BlockSizeShift){
			for(unsigned int i=0;i<N;i++){
				T* Src1=(T*)Temp;
				T* Dest1=(T*)(Dest+(i<<BlockSizeShift));
				for(unsigned int p=0;p<(2UL<<BlockSizeShift);p++) Src1[p]=Dest1[p];
				for(unsigned int p=0;p<(1UL<<BlockSizeShift);p++){
					Dest1[p]=Src1[p<<1];
					Dest1[p+(1UL<<BlockSizeShift)]=Src1[(p<<1)+1];
				}
			}
		}
	};
}

#endif
