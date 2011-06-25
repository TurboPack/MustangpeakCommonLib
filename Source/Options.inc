{$I Compilers.inc}

{$IFDEF COMPILER_7_UP}
  {$DEFINE USETHEMES}
  {$UNDEF JCL_THEME}
{$ENDIF}

{$IFDEF JCL_THEME}
  {$DEFINE USETHEMES}
{$ENDIF}

// -----------------------------------------------------------------------------
// Use unicode-aware inplace editors. This implementation uses Troy Wolbrink's
// TNT controls, which can be found at: www.tntware.com
{.$DEFINE TNTSUPPORT}


