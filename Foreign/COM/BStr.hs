{-# LANGUAGE ForeignFunctionInterface, GeneralizedNewtypeDeriving #-}

module Foreign.COM.BStr (

  BStrPtr(..),
  BStr(..),
  createBStr,
  freeBStr,
  withBStr,
  sizeBytesBStr,
  lengthBStr,
  stringFromBStr

) where

import Control.Exception (bracket)
import Foreign.Ptr
import Foreign.C.String
import Data.Word
import Foreign.Storable

type BStrPtr = CWString
newtype BStr = BStr BStrPtr deriving (Eq, Ord, Storable)

foreign import stdcall "oleauto.h SysAllocStringLen" prim_SysAllocStringLen :: CWString -> Word32 -> IO BStrPtr
foreign import stdcall "oleauto.h SysFreeString"     prim_SysFreeString :: BStrPtr -> IO ()
foreign import stdcall "oleauto.h SysStringLen"      prim_SysStringLen :: BStrPtr -> IO Word32

-- |Create a new @BStr@ from a @String@. The returned BStr must be freed using 'freeBStr'
createBStr :: String -> IO BStr
createBStr s = do
  p <- withCWStringLen s $ \(a,b)-> prim_SysAllocStringLen a $ fromIntegral b
  return $ BStr p

-- |Free a previously created BStr
freeBStr :: BStr -> IO ()
freeBStr (BStr p) = prim_SysFreeString p

-- |Bracket the allocation & freeing of a BStr
withBStr :: String -> (BStr -> IO a) -> IO a
withBStr s = bracket (createBStr s) freeBStr

-- |Size in bytes of a BStr
sizeBytesBStr :: BStr -> IO Int
sizeBytesBStr (BStr p) = peek $ plusPtr p (-4)

-- |Length in characters of a BStr
lengthBStr :: BStr -> IO Int
lengthBStr (BStr p) = prim_SysStringLen p >>= return . fromIntegral

-- |Get the Haskell String from a BStr
stringFromBStr :: BStr -> IO String
stringFromBStr b@(BStr p) = do
  l <- lengthBStr b
  peekCWStringLen (p, l)
  
  
