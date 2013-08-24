-- |Central module re-exporting all of hcom modules
module Foreign.COM (
  module Foreign.COM.COM,
  module Foreign.COM.BStr,
  module Foreign.COM.SafeArray,
  module Foreign.COM.Variant    ) where

import Foreign.COM.COM
import Foreign.COM.BStr
import Foreign.COM.SafeArray
import Foreign.COM.Variant

