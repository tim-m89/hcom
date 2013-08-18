#ifdef __cplusplus
extern "C" {
#endif

#include <Unknwn.h>


void __cdecl unknown_release(IUnknown* obj)
{
  obj->lpVtbl->Release(obj);
}

void __cdecl safearray_destroy(SAFEARRAY* sa)
{
  SafeArrayDestroy(sa);
}


#ifdef __cplusplus
}
#endif


