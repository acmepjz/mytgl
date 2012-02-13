#ifndef MYTMEMORYMANAGEMENT_HPP
#define MYTMEMORYMANAGEMENT_HPP

#include "MyTGL2/MytConfig.hpp"

#include <stdlib.h>
#include <malloc.h>
#include <string.h>

namespace Myt{
	template<class T>
	inline T* Malloc(size_t Size){
		return (T*)malloc(sizeof(T)*Size);
	}

	template<>
	inline void* Malloc<void>(size_t Size){
		return malloc(Size);
	}

	template<class T>
	inline bool Malloc2(T*& Memory,size_t Size){
		Memory=Malloc<T>(Size);
		return Memory!=NULL;
	}

	template<class T>
	inline T* Realloc(T* Memory,size_t NewSize){
		return (T*)realloc(Memory,sizeof(T)*NewSize);
	}

	template<>
	inline void* Realloc<void>(void* Memory,size_t NewSize){
		return realloc(Memory,NewSize);
	}

	template<class T>
	inline bool Realloc2(T*& Memory,size_t NewSize){
		T* tmp=Realloc<T>(Memory,NewSize);
		if(tmp){
			Memory=tmp;
			return true;
		}else{
			return false;
		}
	}

	inline void Free(void* lp){
		free(lp);
	}

	template<class T>
	inline T* AlignedMalloc(size_t Size,size_t Alignment){
#ifdef WIN32
		return (T*)_aligned_malloc(sizeof(T)*Size,Alignment);
#else
		return (T*)memalign(Alignment,sizeof(T)*Size);
#endif
	}

	template<class T,size_t Alignment>
	inline T* AlignedMalloc(size_t Size){
#ifdef WIN32
		return (T*)_aligned_malloc(sizeof(T)*Size,Alignment);
#else
		return (T*)memalign(Alignment,sizeof(T)*Size);
#endif
	}

	inline void AlignedFree(void* lp){
#ifdef WIN32
		_aligned_free(lp);
#else
		free(lp);
#endif
	}
}

#endif
