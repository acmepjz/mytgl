#ifndef MYTGRADIENT_HPP
#define MYTGRADIENT_HPP

#include "MyTGL2/Util/MytMemoryManagement.hpp"
#include "MyTGL2/Util/MytFunctions.hpp"
#include <stdlib.h>
#include <string.h>

namespace Myt{
	template <class T,class T_float>
	class IGradient{
	public:
		virtual T Get(T_float x)=0;
	};
	template <class T,class T_float,class T_Gradient>
	class IGradientImpl:public IGradient<T,T_float>,public T_Gradient{
	public:
		virtual T Get(T_float x){
			return static_cast<T_Gradient*>(this)->Get(x);
		}
	};

	template <class T,class T_float>
	class ConstantGradient{
	public:
		T Value;
	public:
		inline T Get(T_float x){
			return Value;
		}
	};

	template <class T,class T_float>
	class SimpleLinearGradient{
	public:
		T Value[2];
	public:
		inline T Get(T_float x){
			return LinearFunctions<T,T_float>::Lerp(Value[0],Value[1],x);
		}
	};

	template <class T,class T_float>
	class LinearGradient{
	public:
		struct PointType{
			T_float x;
			T y;
			static int QSortCompareFunction(const void* lp1_,const void* lp2_){
				const PointType& lp1=*(const PointType*)lp1_;
				const PointType& lp2=*(const PointType*)lp2_;
				return lp1.x<lp2.x?-1:(lp1.x>lp2.x?1:0);
			}
		};
	protected:
		int Count;
		PointType* Value;
	public:
		inline LinearGradient(){
			Count=0;
			Value=NULL;
		}
		//data in lpSrc should be sorted
		inline LinearGradient(int Count_,const PointType* lpSrc=NULL){
			if(Count_>0){
				Count=Count_;
				Value=Malloc<PointType>(Count_);
				if(lpSrc!=NULL) memcpy(Value,lpSrc,sizeof(PointType)*Count_);
			}else{
				Count=0;
				Value=NULL;
			}
		}
		~LinearGradient(){
			if(Value) Free(Value);
		}
		inline void Destroy(){
			Count=0;
			if(Value) Free(Value);
			Value=NULL;
		}
		inline void Create(int Count_,const PointType* lpSrc=NULL){
			if(Value) Free(Value);
			if(Count_>0){
				Count=Count_;
				Value=Malloc<PointType>(Count_);
				if(lpSrc!=NULL) memcpy(Value,lpSrc,sizeof(PointType)*Count_);
			}else{
				Count=0;
				Value=NULL;
			}
		}
		inline int GetCount() const{
			return Count;
		}
		inline const PointType& operator[](int idx) const{
			return Value[idx];
		}
		inline PointType& operator[](int idx){
			return Value[idx];
		}
		inline void Resize(int Count_){
			if(Count_>=0){
				Count=Count_;
				Value=Realloc(Value,Count_)
			}
		}
		//inserted point will not be sorted
		//not recommended to use this function
		inline void PushBack(const PointType& p){
			Value=Realloc(Value,++Count);
			Value[Count-1]=p;
		}
		//inserted point will be sorted and put to proper location
		//not recommended to use this function
		inline void Insert(const PointType& p){
			Value=Realloc(Value,++Count);
			if(Count<2 || p.x>=Value[Count-2].x){
				Value[Count-1]=p;
			}else{
				//binary search
				int i=0,j=Count-2;
				for(;i<j;){
					int k=(i+j)/2;
					if(p.x>=Value[k].x) i=k+1;
					else j=k-1;
				}
				if(p.x>=Value[i].x) i++;
				for(j=Count-1;j>i;j--) Value[j]=Value[j-1];
				Value[i]=p;
			}
		}
		inline void Sort(){
			if(Count>1) qsort(Value,Count,sizeof(PointType),PointType::QSortCompareFunction);
		}
		//not recommended to use this function
		inline T Get(T_float x){
			if(Count<1) return Constants<T>::Zero();
			else if(Count<2 || x<=Value[0].x) return Value[0].y;
			else if(x>=Value[Count-1].x) return Value[Count-1].y;
			else{
				//binary search
				int i=0,j=Count-1;
				for(;i<j;){
					int k=(i+j)/2;
					if(x>=Value[k].x) i=k+1;
					else j=k-1;
				}
				if(x>=Value[i].x) i++;
				return LinearFunctions<T,T_float>::Lerp(Value[i-1].y,Value[i].y,
					(x-Value[i-1].x)/(Value[i].x-Value[i-1].x));
			}
		}
	};

	template <class T,class T_float>
	class SequentialLinearGradient:public LinearGradient<T,T_float>{
	protected:
		int LastIndex;
	public:
		inline T Get(T_float x){
			if(Count<1) return Constants<T>::Zero();
			else if(Count<2 || x<=Value[0].x) return Value[0].y;
			else if(x>=Value[Count-1].x) return Value[Count-1].y;
			else{
				int i=LastIndex,j;
				do{
					if(i>0 && i<Count){
						if(x>=Value[i-1].x){
							if(x<=Value[i].x){
								break;
							}else if(i<Count-1 && x<=Value[i+1].x){
								LastIndex=++i;
								break;
							}else{
								i++;
								j=Count-1;
							}
						}else if(i>1 && x>=Value[i-2].x){
							LastIndex=--i;
							break;
						}else{
							j=i-2;
							i=0;
						}
					}else{
						i=0;
						j=Count-1;
					}
					//binary search
					for(;i<j;){
						int k=(i+j)/2;
						if(x>=Value[k].x) i=k+1;
						else j=k-1;
					}
					if(x>=Value[i].x) i++;
					LastIndex=i;
				}while(false);
				return LinearFunctions<T,T_float>::Lerp(Value[i-1].y,Value[i].y,
					(x-Value[i-1].x)/(Value[i].x-Value[i-1].x));
			}
		}
	};
}

#endif