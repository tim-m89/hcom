{-# LANGUAGE ForeignFunctionInterface #-}

module Foreign.COM.BStr where

import Control.Exception (bracket)
import Foreign.Ptr
import Foreign.C.String
import Data.Word

type BStr = Ptr ()

foreign import stdcall "oleauto.h SysAllocStringLen" prim_SysAllocStringLen :: CWString -> Word32 -> IO BStr
foreign import stdcall "oleauto.h SysFreeString"     prim_SysFreeString :: BStr -> IO ()

sysAllocString :: String -> IO BStr
sysAllocString s = withCWStringLen s $ \(a,b)-> prim_SysAllocStringLen a $ fromIntegral b

sysFreeString :: BStr -> IO ()
sysFreeString = prim_SysFreeString

withBStr :: String -> (BStr -> IO a) -> IO a
withBStr s = bracket (sysAllocString s) (sysFreeString)

