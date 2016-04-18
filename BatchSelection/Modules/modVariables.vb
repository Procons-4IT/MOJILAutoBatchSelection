Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String

    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public LocalCurrency As String
    Public systemcurrency As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public strDocEntry As String

    Public strImportErrorLog As String = ""
    Public companyStorekey As String = ""
    Public strDatabasename As String = "WMS"
    Public strSKUExportTable As String = "[WMS].dbo.[SKU_Export]"
    Public strSOExportTable As String = "[WMS].dbo.[SO_Export]"
    Public strARCRExportTable As String = "[WMS].dbo.[ARCR_Export]"
    Public strPOExportTable As String = "[WMS].dbo.[PO_Export]"

    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public strItemSelectionQuery As String = ""
    Public frmSourcePaymentform As SAPbouiCOM.Form
    Public frmSourceForm As SAPbouiCOM.Form

    Public intSelectedMatrixrow As Integer = 99999

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum
    Public Const frm_BatchSelect As String = "42"
    Public Const frm_BatchSetup As String = "41"
    Public Const frm_PickBatchSelect As String = "1320000126"


    Public Const frm_ItemType As String = "frm_ItemType"
    Public Const xml_ItemType As String = "frm_ItemType.xml"
    Public Const mnu_ItemType As String = "Z_mnu_PR006"

    Public Const frm_DriverList As String = "frm_DriverList"
    Public Const xml_DriverList As String = "frm_DriverMaster.xml"
    Public Const mnu_DriverList As String = "Z_mnu_PR003"

    Public Const frm_RouteMaster As String = "frm_RouteMaster"
    Public Const xml_RouteMaster As String = "frm_RouteMaster.xml"
    Public Const mnu_RouteMaster As String = "Z_mnu_PR004"

    Public Const frm_CustomerRoute As String = "frm_CutRoute"
    Public Const xml_CustomerRoute As String = "xml_CutRoute.xml"
    Public Const mnu_CustomerRoute As String = "Z_mnu_PR005"


    Public Const frm_BranchMapping As String = "frm_Trans"
    Public Const mnu_Branch As String = "Z_mnu_TR002"
    Public Const xml_BranchMapping As String = "frm_Branch.xml"

    'Public Const frm_BrandMapping As String = "frm_Brand"
    'Public Const mnu_Brand As String = "mnu_Brand"
    'Public Const xml_BrandMapping As String = "frm_Brand.xml"

    Public Const frm_InventoryTransfer As String = "940"
    Public Const frm_InventoryTransferRequest As String = "1250000940"
    Public Const frm_GoodsReceipt As String = "721"
    Public Const frm_GRPO As String = "143"

    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_FuturaSetup As String = "frm_FuturaSetup"
    Public Const frm_StockRequest As String = "frm_StRequest"
    Public Const frm_itemmaster As String = "150"
    Public Const frm_BPMaster As String = "134"
    Public Const frm_InvSO As String = "frm_InvSO"
    Public Const frm_Warehouse As String = "62"
    Public Const frm_SalesOrder As String = "139"
    Public Const frm_ARCreditMemo As String = "179"
    Public Const frm_PurchaseOrder As String = "142"
    Public Const frm_Invoice As String = "133"
    Public Const frm_Import As String = "frm_Import"
    Public Const frm_Export As String = "frm_Export"
    Public Const frm_Futura_Integ As String = "frm_Futura_Integ"

  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_CloseOrderLines As String = "DABT_910"
    Public Const mnu_InvSO As String = "DABT_911"

    Public Const mnu_Import As String = "Z_mnu_FU003"
    Public Const mnu_FutureInt As String = "Z_mnu_FU004"
    'Public Const mnu_Export As String = "Z_mnu_D002"
    'Public Const mnu_Mapping As String = "Z_mnu_FU003"
    
    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_StRequest As String = "StRequest.xml"
    Public Const xml_InvSO As String = "frm_InvSO.xml"
    Public Const xml_Import As String = "frm_Import.xml"
    Public Const xml_Export As String = "frm_Export.xml"
    Public Const xml_Futurasetup As String = "frm_FuturaSetup.xml"
    Public Const xml_FuturaInteg As String = "frm_Futura_Integ.xml"


End Module
