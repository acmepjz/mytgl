#ifndef MYTSTATICASSERT_HPP
#define MYTSTATICASSERT_HPP

namespace Myt{
	/** \brief Simple static assertion.
	\ingroup Utility

	\param b Boolean condition. When b==true there is a specialization does nothing.
	When b==false the template is undefined.
	*/
	template<bool b>
	class StaticAssert;

	/** \brief Simple static assertion (b==true specialization).
	\ingroup Utility
	\sa StaticAssert
	*/
	template<>
	class StaticAssert<true>{
	};
}

#endif