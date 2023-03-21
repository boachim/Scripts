function GetProductDisplayName($skuId,$productNames) 
{
    if ($productNames.ContainsKey($skuId)) {
        return $productNames[$skuId]
    }
    else {
        #License Info file might be out of date, couldn't find new product
        return $skuId
    }
}

function GetServicePlanDisplayName($skuId, $servicePlanId, $servicePlanName) {
    $planKey = $skuId + $servicePlanId
    if ($planNames.ContainsKey($planKey)) {
        return $planNames[$planKey]
    }
    else {
        #License Info file might be out of date, couldn't find new product/plan
        return $servicePlanName
    }
}

Export-ModuleMember -Function GetProductInfo
Export-ModuleMember -Function GetLicenseTypesInTenant
Export-ModuleMember -Function GetProductDisplayName
Export-ModuleMember -Function GetServicePlanDisplayName