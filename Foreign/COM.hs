{-# LANGUAGE ForeignFunctionInterface #-}

module Foreign.COM where

import System.Win32.DLL
import System.Win32.Types
import Control.Monad as M
import Data.Word
import Numeric (showHex)
import Foreign
import Foreign.C
import Control.Exception (bracket)


data GUID = GUID Word32 Word16 Word16 Word8 Word8 Word8 Word8 Word8 Word8 Word8 Word8 
  deriving (Show, Eq)

instance Storable GUID where
  sizeOf    _ = 16
  alignment _ = 4


  peek guidPtr = do
    a  <- peek $ plusPtr guidPtr 0 
    b  <- peek $ plusPtr guidPtr 4
    c  <- peek $ plusPtr guidPtr 6
    d0 <- peek $ plusPtr guidPtr 8
    d1 <- peek $ plusPtr guidPtr 9
    d2 <- peek $ plusPtr guidPtr 10
    d3 <- peek $ plusPtr guidPtr 11
    d4 <- peek $ plusPtr guidPtr 12
    d5 <- peek $ plusPtr guidPtr 13
    d6 <- peek $ plusPtr guidPtr 14
    d7 <- peek $ plusPtr guidPtr 15
    return $ GUID a b c d0 d1 d2 d3 d4 d5 d6 d7

  poke guidPtr (GUID a b c d0 d1 d2 d3 d4 d5 d6 d7) = do
    poke (plusPtr guidPtr 0)  a
    poke (plusPtr guidPtr 4)  b
    poke (plusPtr guidPtr 6)  c
    poke (plusPtr guidPtr 8)  d0
    poke (plusPtr guidPtr 9)  d1
    poke (plusPtr guidPtr 10) d2
    poke (plusPtr guidPtr 11) d3
    poke (plusPtr guidPtr 12) d4
    poke (plusPtr guidPtr 13) d5
    poke (plusPtr guidPtr 14) d6
    poke (plusPtr guidPtr 15) d7


type IID = GUID
type REFIID = Ptr IID
type CLSID = GUID
type REFCLSID = Ptr CLSID
type LIBID = GUID
type REFLIBID = Ptr LIBID
type CATID = GUID
type REFCATID = Ptr CATID


foreign import stdcall "OleAuto.h LoadTypeLibEx"    loadTypeLibEx    :: LPCWSTR -> Int -> Ptr IntPtr -> IO HRESULT
foreign import stdcall "OleAuto.h LoadRegTypeLib"   loadRegTypeLib   :: REFLIBID -> WORD -> WORD -> LCID ->  Ptr IntPtr -> IO HRESULT
foreign import stdcall "Objbase.h CoInitializeEx"   coInitializeEx   :: Ptr () -> Int32 -> IO HRESULT
foreign import stdcall "Objbase.h CoUninitialize"   coUninitialize   :: IO ()
foreign import stdcall "Objbase.h CoCreateInstance" coCreateInstance :: REFCLSID -> Ptr () -> DWORD -> REFIID -> Ptr (Ptr IntPtr) -> IO HRESULT


type ComObjectInternal = Ptr (FunPtr ())
type ComObject = Ptr ComObjectInternal
type ComObjectAuto = ForeignPtr ComObjectInternal

type QueryInterface = ComObject -> REFIID -> Ptr ComObject -> IO HRESULT
foreign import stdcall "dynamic" makeQueryInterface :: FunPtr QueryInterface -> QueryInterface
iUnknown_queryInterface = (0::Int, makeQueryInterface)

type AddRef = ComObject -> IO LONG
foreign import stdcall "dynamic" makeAddRef :: FunPtr AddRef -> AddRef
iUnknown_addRef = (1::Int, makeAddRef)

type Release = ComObject -> IO LONG
foreign import stdcall "dynamic" makeRelease :: FunPtr Release -> Release
iUnknown_release = (2::Int, makeRelease)

foreign import ccall "hcom.c &unknown_release" iunknown_release :: FunPtr (ComObject -> IO ())
comObjectAuto :: ComObject -> IO ComObjectAuto
comObjectAuto obj = do
  fp <- newForeignPtr (iunknown_release) obj
  return fp

withComObjectAuto :: ComObjectAuto -> (ComObject -> IO a) -> IO a
withComObjectAuto = withForeignPtr

getComFunc :: ComObject -> (Int, (FunPtr a -> b)) -> IO b
getComFunc this (index, makeFun) = do
    funPtr <- peek this >>= flip peekElemOff index
    return $ makeFun $ castFunPtr funPtr

comObjectFunc :: ComObject -> (Int, FunPtr a -> ComObject -> b) -> IO b
comObjectFunc this f = do
  func <- getComFunc this f
  return $ func this

