{-# LANGUAGE GADTs #-}

module Foreign.COM.Variant where

import Data.Word
import Foreign.Ptr
import Foreign.Storable

data Variant = Variant VarType Word64 deriving (Show, Eq)
type VarType = Word16

getVarType :: Variant -> VarType
getVarType (Variant vt vd) = vt

vt_EMPTY	= 0 :: VarType
vt_NULL	= 1 :: VarType
vt_I2	= 2 :: VarType
vt_I4	= 3 :: VarType
vt_R4	= 4 :: VarType
vt_R8	= 5 :: VarType
vt_CY	= 6 :: VarType
vt_DATE	= 7 :: VarType
vt_BSTR	= 8 :: VarType
vt_DISPATCH	= 9 :: VarType
vt_ERROR	= 10 :: VarType
vt_BOOL	= 11 :: VarType
vt_VARIANT	= 12 :: VarType
vt_UNKNOWN	= 13 :: VarType
vt_DECIMAL	= 14 :: VarType
vt_I1	= 16 :: VarType
vt_UI1	= 17 :: VarType
vt_UI2	= 18 :: VarType
vt_UI4	= 19 :: VarType
vt_I8	= 20 :: VarType
vt_UI8	= 21 :: VarType
vt_INT	= 22 :: VarType
vt_UINT	= 23 :: VarType
vt_VOID	= 24 :: VarType
vt_HRESULT	= 25 :: VarType
vt_PTR	= 26 :: VarType
vt_SAFEARRAY	= 27 :: VarType
vt_CARRAY	= 28 :: VarType
vt_USERDEFINED	= 29 :: VarType
vt_LPSTR	= 30 :: VarType
vt_LPWSTR	= 31 :: VarType
vt_RECORD	= 36 :: VarType
vt_INT_PTR	= 37 :: VarType
vt_UINT_PTR	= 38 :: VarType
vt_FILETIME	= 64 :: VarType
vt_BLOB	= 65 :: VarType
vt_STREAM	= 66 :: VarType
vt_STORAGE	= 67 :: VarType
vt_STREAMED_OBJECT	= 68 :: VarType
vt_STORED_OBJECT	= 69 :: VarType
vt_BLOB_OBJECT	= 70 :: VarType
vt_CF	= 71 :: VarType
vt_CLSID	= 72 :: VarType
vt_VERSIONED_STREAM	= 73 :: VarType
vt_BSTR_BLOB	= 0xfff :: VarType
vt_VECTOR	= 0x1000 :: VarType
vt_ARRAY	= 0x2000 :: VarType
vt_BYREF	= 0x4000 :: VarType
vt_RESERVED	= 0x8000 :: VarType
vt_ILLEGAL	= 0xffff :: VarType
vt_ILLEGALMASKED	= 0xfff :: VarType
vt_TYPEMASK	= 0xfff :: VarType

instance Storable Variant where
    sizeOf    _ = 16
    alignment _ = 8

    peek ptr = do
        a  <- peek $ plusPtr ptr 0 
        b  <- peek $ plusPtr ptr 8
        return $ Variant a b

    poke ptr (Variant a b) = do
        poke (plusPtr ptr 0) a
        poke (plusPtr ptr 2) (0 :: Word16)
        poke (plusPtr ptr 4) (0 :: Word16)
        poke (plusPtr ptr 6) (0 :: Word16)
        poke (plusPtr ptr 8) b


type VariantBool = Word16
varTrue = 0xFFFF :: VariantBool
varFalse = 0x0000 :: VariantBool

class HVarType a where
  hVarType :: a -> VarType

instance HVarType Int where
  hVarType _ = vt_INT

