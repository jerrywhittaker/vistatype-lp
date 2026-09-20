"""A small stand-in for src/ribbon/customUI14.xml.

Small enough to reason about, shaped like the real one: two tabs a transcriber should be
able to hide and reorder, the hidden QAT tab the installed toolbars depend on, and the
retired-buttons tab.
"""

RIBBON = """<?xml version="1.0" encoding="utf-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tab_LP" label="VistaType LP">
        <group id="grp_Files" label="Files" imageMso="FileSave">
          <button id="btn_Lp_One" label="One" imageMso="HappyFace"/>
          <button id="btn_Lp_Two" label="Two &amp;&amp; More" imageMso="SadFace"/>
          <control idMso="Bold"/>
        </group>
      </tab>
      <tab id="tab_BRL" label="Braille Macros">
        <group id="grp_Brl" label="Braille">
          <button id="btn_Dx_One" label="One" imageMso="Cat"/>
        </group>
      </tab>
      <tab id="tab_LP_and_BRL_QAT_Icons" label="QAT Icons" visible="false">
        <group id="grp_Qat" label="Qat">
          <button id="btn_Lp_One" label="One" imageMso="HappyFace"/>
          <button id="btn_Dx_One" label="One" imageMso="Cat"/>
        </group>
      </tab>
      <tab id="tab_VT_Retired_Buttons" label="Retired" visible="false">
        <group id="grp_Ret" label="Retired">
          <button id="btn_Lp_Gone" label="Gone" imageMso="Dog"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""
