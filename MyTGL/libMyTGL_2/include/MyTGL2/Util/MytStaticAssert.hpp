#ifndef MYTSTATICASSERT_HPP
#define MYTSTATICASSERT_HPP

namespace Myt{
	template<bool b>
	class StaticAssert;

	template<>
	class StaticAssert<true>{
	};

	template<>
	class StaticAssert<false>{
	private:
		inline StaticAssert();
	};
}

#endif