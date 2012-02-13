#ifndef MYTFFT_HPP
#define MYTFFT_HPP

#include "MyTGL2/Util/MytMemoryManagement.hpp"
#include "MyTGL2/Util/MytFunctions.hpp"
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
		static void Calc(const Complex<T>* Src,Complex<T>* Dest,unsigned int N,bool Inverse,bool Normalize){
			Shuffle(Src,Dest,N);
			CalcInternalAndPostProcess(Dest,N,Inverse,Normalize);
		}
		static void Calc(Complex<T>* Dest,unsigned int N,bool Inverse,bool Normalize){
			Shuffle(Dest,N);
			CalcInternalAndPostProcess(Dest,N,Inverse,Normalize);
		}
		static void Calc(const Complex<T>* Src,Complex<T>* Dest,bool Inverse,bool Normalize,const FFTShuffleProvider& obj){
			const unsigned int N=obj.N;
			Shuffle(Src,Dest,N,obj);
			CalcInternalAndPostProcess(Dest,N,Inverse,Normalize);
		}
		static void Calc(Complex<T>* Dest,bool Inverse,bool Normalize,const FFTShuffleProvider& obj){
			const unsigned int N=obj.N;
			Shuffle(Dest,N,obj);
			CalcInternalAndPostProcess(Dest,N,Inverse,Normalize);
		}
	private:
		static inline void CalcInternalAndPostProcess(Complex<T>* Dest,unsigned int N,bool Inverse,bool Normalize){
			CalcInternal(Dest,N,Inverse);
			if(Normalize){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}else if(Inverse){
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
		static void CalcReal(const T* Src,Complex<T>* Dest,unsigned int N,bool Normalize){
			Shuffle((const Complex<T>*)Src,Dest,N>>1);
			CalcRealInternalAndPostProcess(Dest,N>>1,Normalize);
		}
		//N must be even, power-of-two
		static void CalcReal(T* Dest,unsigned int N,bool Normalize){
			Shuffle((Complex<T>*)Dest,N>>1);
			CalcRealInternalAndPostProcess((Complex<T>*)Dest,N>>1,Normalize);
		}
		//real data size should be 2*obj.N
		static void CalcReal(const T* Src,Complex<T>* Dest,bool Normalize,const FFTShuffleProvider& obj){
			Shuffle((const Complex<T>*)Src,Dest,obj.N,obj);
			CalcRealInternalAndPostProcess(Dest,obj.N,Normalize);
		}
		//real data size should be 2*obj.N
		static void CalcReal(T* Dest,bool Normalize,const FFTShuffleProvider& obj){
			Shuffle((Complex<T>*)Dest,obj.N,obj);
			CalcRealInternalAndPostProcess((Complex<T>*)Dest,obj.N,Normalize);
		}
	private:
		static inline void CalcRealInternalAndPostProcess(Complex<T>* Dest,unsigned int N,bool Normalize){
			CalcInternal(Dest,N,false);
			//post process
			Dest[0]=Complex<T>::Make(
				Dest[0].Re+Dest[0].Im, //Dest[0]
				Dest[0].Re-Dest[0].Im  //Dest[N>>1]
				);

			const T_float pi=T_float(-3.14159265358979323846);
			const T_float delta=pi/T_float(N);
			const Complex<T_float> Multiplier={cos(delta),sin(delta)};
			Complex<T_float> Factor={0,-1};

			for(unsigned int k=1;k<=(N>>1);k++){
				Factor*=Multiplier;

				Complex<T> tmp[1];
				tmp[0]=!Dest[N-k];

				tmp[1]=Dest[k]+tmp[0];
				tmp[0]=(Dest[k]-tmp[0])*Factor;

				Dest[k]=(tmp[1]+tmp[0])*T_float(0.5);
				Dest[N-k].Re=(tmp[1].Re-tmp[0].Re)*T_float(0.5);
				Dest[N-k].Im=(tmp[0].Im-tmp[1].Im)*T_float(0.5);
			}
			//over
			if(Normalize){
				T_float f=T_float(1.0)/sqrt(T_float(N<<1));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}
		}
	public:
		static void CalcND(const Complex<T>* Src,Complex<T>* Dest,unsigned int Dimension,const FFTShuffleProvider* obj,bool Inverse,bool Normalize){
			//calc size
			unsigned int N=1UL;
			for(unsigned int d=0;d<Dimension;d++) N<<=obj[d].Shift;
			//calc d=0
			unsigned int Step=obj[0].N,BlockSizeShift=obj[0].Shift;
			for(unsigned int i=0;i<N;i+=Step){
				Shuffle(Src+i,Dest+i,Step,obj[0]);
				CalcInternal(Dest+i,Step,Inverse);
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
			if(Normalize){
				T_float f=T_float(1.0)/sqrt(T_float(N));
				for(unsigned int i=0;i<N;i++) Dest[i]*=f;
			}else if(Inverse){
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
	};
}

#endif
