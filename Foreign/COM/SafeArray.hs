{-# LANGUAGE ForeignFunctionInterface, ScopedTypeVariables, GeneralizedNewtypeDeriving #-}

module Foreign.COM.SafeArray (

  SafeArrayPtr(..),
  SafeArray(..),
  safeArrayCreateVector,
  safeArrayLength,
  safeArraySetData,
  arrayToSafeArray,
  safeArrayToArray,
  withSafeArray,

  -- * Explicit VarType
  -- $vT

  withSafeArrayVT,
  arrayToSafeArrayVT,
  safeArrayCreateVectorVT

) where

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

type SafeArrayPtr a = Ptr a

newtype SafeArray a = SafeArray (SafeArrayPtr a) deriving (Eq, Ord, Storable)

instance HVarType (SafeArray a) where
  hVarType _ = vt_SAFEARRAY

foreign import stdcall "oleauto.h SafeArrayCreateVector" prim_SafeArrayCreateVector :: VarType -> Int32 -> Word32 -> IO (SafeArrayPtr a)
foreign import stdcall "oleauto.h SafeArrayAccessData"   prim_SafeArrayAccessData   :: SafeArrayPtr a -> Ptr (Ptr a) -> IO HRESULT
foreign import stdcall "oleauto.h SafeArrayUnaccessData" prim_SafeArrayUnaccessData :: SafeArrayPtr a -> IO HRESULT
foreign import stdcall "oleauto.h SafeArrayDestroy"      prim_SafeArrayDestroy      :: SafeArrayPtr a -> IO HRESULT
foreign import ccall   "hcom.c &safearray_destroy"       primc_safeArrayDestroy     :: FunPtr (SafeArrayPtr a -> IO ())
foreign import stdcall "oleauto.h SafeArrayGetLBound"    prim_SafeArrayGetLBound    :: SafeArrayPtr a -> Word32 -> Ptr Int32 -> IO HRESULT
foreign import stdcall "oleauto.h SafeArrayGetUBound"    prim_SafeArrayGetUBound    :: SafeArrayPtr a -> Word32 -> Ptr Int32 -> IO HRESULT


safeArrayCreateVector :: forall a. (HVarType a) => Int -> IO (SafeArray a)
safeArrayCreateVector z = do
  let vt = hVarType (undefined::a)
  safeArrayCreateVectorVT vt z

safeArrayLength :: SafeArray a -> IO Int
safeArrayLength (SafeArray sap) = alloca $ \pu-> alloca $ \pl-> do
  prim_SafeArrayGetUBound sap 1 pu
  prim_SafeArrayGetLBound sap 1 pl
  arrayUpper <- peek pu
  arrayLower <- peek pl
  return $ fromIntegral (arrayUpper - arrayLower + 1)

safeArraySetData :: (Storable a) => SafeArray a -> [a] -> IO ()
safeArraySetData (SafeArray sap) x = alloca $ \p-> do
  prim_SafeArrayAccessData sap p
  p2 <- peek p
  pokeArray p2 x
  prim_SafeArrayUnaccessData sap
  return ()


arrayToSafeArray :: forall a. (Storable a, HVarType a) => [a] -> IO (SafeArray a)
arrayToSafeArray x = let vt = hVarType (undefined::a) in arrayToSafeArrayVT vt x 

safeArrayToArray :: (Storable a) => SafeArray a -> IO [a]
safeArrayToArray sa@(SafeArray sap) = alloca $ \p-> do
  l <- safeArrayLength sa
  prim_SafeArrayAccessData sap p
  p2 <- peek p
  x <- peekArray l p2
  prim_SafeArrayUnaccessData sap
  return x

withSafeArray :: forall a b. (Storable a, HVarType a) => [a] -> (SafeArray a -> IO b) -> IO b
withSafeArray x f = let vt = hVarType (undefined::a) in withSafeArrayVT vt x f


{- $vT
    Variantions of the above that allow the caller to explicitly
    choose the VarType rather than have it inferred from the type
-}


withSafeArrayVT :: (Storable a) => VarType -> [a] -> (SafeArray a -> IO b) -> IO b
withSafeArrayVT vt x f = bracket (arrayToSafeArrayVT vt x) (\(SafeArray p)-> prim_SafeArrayDestroy p) f

arrayToSafeArrayVT :: (Storable a) => VarType -> [a] -> IO (SafeArray a)
arrayToSafeArrayVT vt x = do
  sa <- safeArrayCreateVectorVT vt $ length x
  safeArraySetData sa x
  return sa

safeArrayCreateVectorVT :: VarType -> Int -> IO (SafeArray a)
safeArrayCreateVectorVT vt z = do
  sap <- prim_SafeArrayCreateVector vt 0 (fromIntegral z)
  return $ SafeArray sap

