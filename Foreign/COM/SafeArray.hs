{-# LANGUAGE ForeignFunctionInterface, ScopedTypeVariables #-}

module Foreign.COM.SafeArray where

import Control.Exception (bracket)
import Data.Int
import Data.Word
import Foreign.COM.Variant
import Foreign.Ptr
import Foreign.Marshal.Alloc
import Foreign.Marshal.Array
import Foreign.Storable
import Foreign.ForeignPtr
import System.Win32

type SafeArray a = Ptr a

foreign import stdcall "oleauto.h SafeArrayCreateVector" prim_SafeArrayCreateVector :: VarType -> Int32 -> Word32 -> IO (SafeArray a)
foreign import stdcall "oleauto.h SafeArrayAccessData"   prim_SafeArrayAccessData   :: SafeArray a -> Ptr (Ptr a) -> IO HRESULT
foreign import stdcall "oleauto.h SafeArrayUnaccessData" prim_SafeArrayUnaccessData :: SafeArray a -> IO HRESULT
foreign import stdcall "oleauto.h SafeArrayDestroy"      prim_SafeArrayDestroy      :: SafeArray a -> IO HRESULT
foreign import ccall   "hcom.c &safearray_destroy"       primc_safeArrayDestroy     :: FunPtr (SafeArray a -> IO ())
foreign import stdcall "oleauto.h SafeArrayGetLBound"    prim_SafeArrayGetLBound    :: SafeArray a -> Word32 -> Ptr Int32 -> IO HRESULT
foreign import stdcall "oleauto.h SafeArrayGetUBound"    prim_SafeArrayGetUBound    :: SafeArray a -> Word32 -> Ptr Int32 -> IO HRESULT

safeArrayCreateVectorVT :: VarType -> Int32 -> Word32 -> IO (SafeArray a)
safeArrayCreateVectorVT = prim_SafeArrayCreateVector

safeArrayCreateVector :: forall a. (HVarType a) => Int32 -> Word32 -> IO (SafeArray a)
safeArrayCreateVector y z = do
  let vt = hVarType (undefined::a)
  safeArrayCreateVectorVT vt y z

safeArrayLength :: SafeArray a -> IO Int
safeArrayLength sa = alloca $ \pu-> alloca $ \pl-> do
  prim_SafeArrayGetUBound sa 1 pu
  prim_SafeArrayGetLBound sa 1 pl
  arrayUpper <- peek pu
  arrayLower <- peek pl
  return $ fromIntegral (arrayUpper - arrayLower + 1)

safeArraySetData :: (Storable a) => SafeArray a -> [a] -> IO ()
safeArraySetData sa x = alloca $ \p-> do
  prim_SafeArrayAccessData sa p
  p2 <- peek p
  pokeArray p2 x
  prim_SafeArrayUnaccessData sa
  return ()

arrayToSafeArrayVT :: (Storable a) => VarType -> [a] -> IO (SafeArray a)
arrayToSafeArrayVT vt x = do
  sa <- safeArrayCreateVectorVT vt 0 $ fromIntegral $ length x
  safeArraySetData sa x
  return sa

arrayToSafeArray :: forall a. (Storable a, HVarType a) => [a] -> IO (SafeArray a)
arrayToSafeArray x = let vt = hVarType (undefined::a) in arrayToSafeArrayVT vt x 

safeArrayToArray :: (Storable a) => SafeArray a -> IO [a]
safeArrayToArray sa = alloca $ \p-> do
  l <- safeArrayLength sa
  prim_SafeArrayAccessData sa p
  p2 <- peek p
  x <- peekArray l p2
  prim_SafeArrayUnaccessData sa
  return x

withSafeArrayVT :: (Storable a) => VarType -> [a] -> (SafeArray a -> IO b) -> IO b
withSafeArrayVT vt x f = bracket (arrayToSafeArrayVT vt x) prim_SafeArrayDestroy f

withSafeArray :: forall a b. (Storable a, HVarType a) => [a] -> (SafeArray a -> IO b) -> IO b
withSafeArray x f = let vt = hVarType (undefined::a) in withSafeArrayVT vt x f

