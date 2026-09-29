$ErrorActionPreference = 'Stop'
$qaTask = Join-Path $PSScriptRoot 'qa'
$xsdTask = Join-Path (Split-Path $PSScriptRoot) 'output/CBI_XML_Macro/Fund Admin - XSD v1.0.xsd'
$nsTask = 'http://www.centralbank.ie/Schemas/AML/2025/A12'
$cacheTask = New-Object -ComObject MSXML2.XMLSchemaCache.6.0
$schemaTask = New-Object -ComObject MSXML2.DOMDocument.6.0
$schemaTask.async = $false
$schemaTask.resolveExternals = $false
$schemaTask.setProperty('ProhibitDTD', $true)
if (-not $schemaTask.load($xsdTask)) { throw $schemaTask.parseError.reason }
$cacheTask.add($nsTask, $schemaTask)
foreach ($nameTask in @('completed_corrected_test.xml','blank_corrected_test.xml')) {
    $domTask = New-Object -ComObject MSXML2.DOMDocument.6.0
    $domTask.async = $false
    $domTask.validateOnParse = $false
    $domTask.schemas = $cacheTask
    if (-not $domTask.load((Join-Path $qaTask $nameTask))) { throw $domTask.parseError.reason }
    $resultTask = $domTask.validate()
    if ($resultTask.errorCode -ne 0) { throw $resultTask.reason }
    "PASS MSXML6: $nameTask"
    $domTask.setProperty('SelectionNamespaces', "xmlns:a='$nsTask'")
    $nodeTask = $domTask.selectSingleNode("//a:A12031[@Identifier='LBV']")
    $nodeTask.setAttribute('EnumerationBWRAConducted','yES')
    if ($domTask.validate().errorCode -eq 0) { throw 'Invalid enumeration accepted' }
    if ($domTask.validateNode($nodeTask).errorCode -eq 0) { throw 'Invalid node accepted' }
    $nodeTask.setAttribute('EnumerationBWRAConducted','Yes')
    $nodeTask = $domTask.selectSingleNode('//a:A12021')
    $savedTask = $nodeTask.getAttribute('LBD')
    $nodeTask.setAttribute('LBD','2147483648')
    if ($domTask.validate().errorCode -eq 0) { throw 'Integer overflow accepted' }
    $nodeTask.setAttribute('LBD', $savedTask)
    $nodeTask = $domTask.selectSingleNode("//a:A12031[@Identifier='LBX']")
    $savedTask = $nodeTask.getAttribute('DateValue')
    $nodeTask.setAttribute('DateValue','2026-02-30')
    if ($domTask.validate().errorCode -eq 0) { throw 'Invalid date accepted' }
    $nodeTask.setAttribute('DateValue', $savedTask)
    $nodeTask = $domTask.selectSingleNode("//a:A12031[@Identifier='LNX']")
    $savedTask = $nodeTask.getAttribute('PercentageValue')
    $nodeTask.setAttribute('PercentageValue','4.7')
    if ($domTask.validate().errorCode -eq 0) { throw 'Percentage overflow accepted' }
    $nodeTask.setAttribute('PercentageValue','0.0000001')
    if ($domTask.validate().errorCode -eq 0) { throw 'Excessive precision accepted' }
    $nodeTask.setAttribute('PercentageValue', $savedTask)
    if ($domTask.validate().errorCode -ne 0) { throw 'Unexpected final validation failure' }
    'PASS MSXML6 negative tests: enum, integer bound, date, percent, precision, source-node validation'
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($domTask)
}
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cacheTask)
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($schemaTask)
