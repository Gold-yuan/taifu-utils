package sichuan.ytf.translate;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

public class Translation {
    static Map<String, String> map = new HashMap<>();

    @Test
    public void test() {
        
        String cn2 = "统一信用代码(身份证号码),许可内容,企业名称,日常监督管理机构,日常监督管理人员,许可证编号,有效期至,许可状态,许可证名称,企业负责人,法定代表人,生产地址,住所,发证日期,发证机关,许可项目,签发人,许可时间,质量负责人";

        
        
        StringBuffer sba = new StringBuffer();
        sba.append("科室地址,生产（经营）地址,许可证号,日常监督管理人员,许可范围,有效期起,许可证有效性,注册地址,有效期至,许可证类别及科室名称,日常监管人员,企业负责人,投诉举报电话,生产地址,食品类别,地址,有效期止,主体类型,经营者名称,医疗机构类别,日常监督管理机构,许可证编号,分类码,社会信用代码,法定代表人(负责任人),发证机关,医疗机构名称,日常监管机构,生产（经营）范围,备注,经济性质,外设仓库地址,住所,发证日期,主体业态,签发人,质量负责人,经营类别,生产地址和范围,经营场所,行政区域,法定代表人(负责人),企业名称,科室负责人,经营项目,法定代表人,GMP证书,产品名称,负责人,状态说明,生产者名称");
        StringBuffer sbb = new StringBuffer();
        sbb.append("departmentAddress,productionAddress,licenseNumber,dailySupervisionManagementPersonnel,licenseScope,validityLicenseStart,validityLicense,registrationAddress,expirationDate,licenseCategoryDepartmentName,dailySupervisor,enterprisePersonCharge,complaintReportingTelephone,productionAddress,foodCategory,address,validityLicenseEnd,mainBodyType,operatorName,medicalInstitutionCategory,dailySupervisionManagementOrganization,licenseNumber,classificationCode,socialCreditCode,legalRepresentative,licensingAuthority,medicalInstitutionName,dailySupervisoryAuthority,productionScope,notes,economicNature,externalWarehouseAddress,residence,issuingDate,mainBusinessForm,issuingPersion,qualityPersonCharge,businessCategory,productionAddressScope,businessPlace,administrativeRegion,legalRepresentative,enterpriseName,departmentHead,businessProject,legalRepresentative,GMPCertificate,productName,personCharge,statusDescription,producerName");
        
        
        
        

        sba.append(",xxxxx");
        sbb.append(",xxxxx");
        sba.append(",科（室）负责人,许可证类别及科（室）名称,科（室）地址");
        sbb.append(",departmentChargePerson,licenseCatetoryAndDepartmentName,departmentAddress");
        sba.append(",企业名称（英文）,认证范围（英文）");
        sbb.append(",enterpriseNameEn,authorityScopeEn");
        sba.append(",生产地址与生产范围,配制地址与配置范围");
        sbb.append(",productionAddressAndScope,configAddressAndscope");
        sba.append(",统一信用代码(身份证号码),许可内容,许可状态,许可证名称,许可时间");
        sbb.append(",socialCreditCode,licenseContext,licenseStatus,licenseName,licenseDate");
        sba.append(",批文时间,变更说明,批文号,文件名称,定点产品");
        sbb.append(",approvalTime,changeInstruction,approvalNumber,documentName,designatedProduct");
        sba.append(",现生产地址,现注册地址,类型");
        sbb.append(",productionAddress,registerAddress,category");
        sba.append(",证书种类,生产类型,生产批件号");
        sbb.append(",certCategory,productionType,productionBatchNumber");
        sba.append(",设备名称,申请单位,发射功率,占用带宽,设备型号,频率容限,频率范围,杂散发射限制,核准证编号");
        sbb.append(",equipmentName,applicantUnit,transmittingPower,occupiedbBandwidth,equipmentType,frequencyTolerance,frequencyRange,miscellaneousEmissionLimitation,approvalCertificateNumber");
        sba.append(",注册商标,批准产量(吨),标志编号,有效期区间");
        sbb.append(",registeredTrademark,approvedOutput,markNumber,validityPeriodInterval");
        sba.append(",网址,认证人,联系人,电话,企业地址第二备用,新邮地址,企业地址备用,邮箱,企业信息码,联系人电话");
        sbb.append(",webUrl,issuingPersion,enterpriseContactPerson,phone,addressTwo,newEmailAddress,addressThree,email,enterpriseInformationCode,contactPersonPhone");
        sba.append(",医疗单位,主治功能或适应症");
        sbb.append(",medicalUnit,primaryFunctionOrIndication");
        sba.append(",有效日期起,有效日期止,GSP许可证有效日期起,GSP许可证有效日期止");
        sbb.append(",validDateBegin,validDateEnd,GSPvalidDateBegin,GSPvalidDateEnd");
        sba.append(",法人 代表,开办企业类别,许可证(备案)号");
        sbb.append(",corporater,startUpBusinessCategory,licenseNumber");
        sba.append(",产品英文名称,产品注册号,注册类别,产品使用范围");
        sbb.append(",productNameEn,productRegisterNo,registerType,productUsedScope");
        sba.append(",药包材名称,使用企业,使用企业所在地区,产品质量标准名称,注册编号");
        sbb.append(",pharmaceuticalPackageName,useEnterpry,useEnterpryArea,productQualityStandardName,registrationNumber");
        sba.append(",英文\\拉丁名,是否处方药,适应症或主治功能,药用辅料名称");
        sbb.append(",nameEn,prescribeDrug,primaryFunction,medicinalExcipientName");
        sba.append(",卫产许可证号,所在地区");
        sbb.append(",healthProductionLicenseNumber,country");
        sba.append(",商品名称,汉语拼英,商标,有效期至止,功效成分,适宜人群,储藏方式,注意事项,用法用量");
        sbb.append(",commodityName,chineseSpelling,trademark,expirationDate,efficacyIngredients,suitablePeople,storageMethod,mattersNeedingAttention,usageDosage");
        sba.append(",企业所在地区,生产品种,QS证号");
        sbb.append(",enterpriseLocation,productionVarieties,QSCertificateNumber");
        sba.append(",登载日期,委托方,受托方");
        sbb.append(",launchDate,principal,trustee");
        sba.append(",企业类别,营业执照号,质量代表人,ip地址");
        sbb.append(",enterpriseType,businessLicenseNo,qualityRepresentative,ipAddress");
        sba.append(",生产加工者名称,生产加工地址");
        sbb.append(",processorManufacturerName,processingManufacturingAddress");
        sba.append(",类型类别,配置范围,经营者(负责人)姓名");
        sbb.append(",typeCategory,configurationScope,managerName");
        sba.append(",科室名称,许可证类别");
        sbb.append(",departmentName,licenceCategory");
        sba.append(",范围,非法人变更为法人");
        sbb.append(",scpoe,changeCorporater");
        sba.append(",英文地址,英文范围,英文发证机关");
        sbb.append(",addressEn,scopeEn,licensingAuthorityEn");
        sba.append(",证书名称,类别,变更(1)");
        sbb.append(",certName,category,change");
        sba.append(",业务编码,销售类型");
        sbb.append(",businessNo,saleType");
        sba.append(",生产企业名称,准产注册证号,首次重新注册,重新注册品种原注册证号");
        sbb.append(",productionEnterpriseName,quasiProductionRegistrationCertificateNumber,firstReregistration,reregistrationVarietiesOriginalRegistrationCertificateNumber");
        sba.append(",产品备案号,是否受托生产");
        sbb.append(",productRecordNumber,fiduciaryProduction");
        sba.append(",服务范围,省辖市,生产类别");
        sbb.append(",serviceScope,provincialMunicipality,productCategory");
        sba.append(",签 发 人,机构地址,主服务器域名,主服务器地址,经营性质,网站负责人");
        sbb.append(",issuer,institutionalAddress,mainServerDomainName,mainServerAddress,businessNature,webSiteManager");
        sba.append(",登记证编号,小作坊名称,产品品种,登记日期,登记机关");
        sbb.append(",registrationCertificateNumber,smallWorkshopName,productVariety,registrationDate,registrationAuthority");
        sba.append(",食品食品添加剂类别,产品类别编号,产品类别名称,产品明细");
        sbb.append(",foodAdditivesCategory,productCategoryNumber,productCategoryName,productDetails");
        sba.append(",食品、食品添加剂类别,类别编号,类别名称,品种明细,产品名称及执行标准");
        sbb.append(",category,categoryNumber,categoryName,categoryDetail,productNameAndExecutiveStandard");
        sba.append(",原证书编号");
        sbb.append(",oldCertificateNumber");
        sba.append(",GSP证书号,GSP证书有效期起至时间");
        sbb.append(",GSPCertNo,GSPCertValidityPeriod");
        sba.append(",是否是体外诊断试剂,产品分类名称,日期");
        sbb.append(",vitroDiagnosticReagent,productClassificationName,date");
        sba.append(",汉语拼音,英文名/拉丁名,曾用名,原批准单位,注册分类,包装,批准文号有效期,包装标签,药品说明书");
        sbb.append(",chinesePinyin,englishNameOrLatinName,usedName,formerApprovalUnit,registrationClassification,packaging,validityPeriodApprovalNumber,packagingLabel,pharmaceuticalInstructions");
        sba.append(",生产国（地区）,生产企业（中文）,生产企业（英文）,在华申报责任单位,在华责任单位地址");
        sbb.append(",productCountry,productionEnterpriseCn,productionEnterpriseEn,chinaDeclareResponsibleUnits,ChinaResponsibleUnitsAddress");
        sba.append(",生产企业地址,批件状态,批件有效期,卫生许可证号,产品名称备注,产品技术要求");
        sbb.append(",productionEnterprisesAddress,batchStatu,batchvValidityPeriodDate,hygieneLicenseNumber,productNameAnnotation,productsTechnicalRequirement");
        sba.append(",经营范围(2002分类),经营范围(2017分类)");
        sbb.append(",operationScope2002,operationScope2017");
        sba.append(",生产产品登记表");
        sbb.append(",productRegistrationForm");
        sba.append(",产品名称或产品分类名称,型号/规格或包装规格,产品描述或主要组成成份");
        sbb.append(",productNameOrProductTypeName,modelSpecificationOrPackageSpecification,productDescriptionOrMainComponents");
        sba.append(",生产厂商名称（英文）,生产厂地址（中文）,注册代理");
        sbb.append(",manufacturerNameEn,manufacturerAddressCn,registerProxy");
        sba.append(",生产国或地区（英文）,生产厂商名称（中文）,生产国或地区（中文）,售后服务机构");
        sbb.append(",countryOrRegionEn,manufacturerNameCn,countryOrRegionCn,aftersalesServiceOrganization");
        sba.append(",型号、规格,结构及组成,适用范围,产品标准,邮编,主要组成成分（体外诊断试剂）,预期用途（体外诊断试剂）,产品储存条件及有效期（体外诊断试剂）");
        sbb.append(",modelSpecification,structureComposition,applicationScope,productStandard,zipCode,mainComponents,expectedUse,storageConditionsAndExpiryDate");
        sba.append(",保护品种编号,药品名称,公告号,保护级别,保护起始日,保护终止日,保护期限");
        sbb.append(",protectionSpeciesNumber,drugName,bulletinNumber,protectionLevel,protectionStartDate,protectionTerminationDate,protectionPeriod");
        sba.append(",品种名称,产品来源,法人代表");
        sbb.append(",varietiesName,productSource,corporater");
        sba.append(",省市,批准延续日期,有效期延续至,批准延续的认证范围,认证GMP版本");
        sbb.append(",provinces,approvalExtensionDate,validityPeriodExtendDate,approvalExtensionCertificationScope,certificationGMPVersion");
        sba.append(",社会信用代码/组织机构代码");
        sbb.append(",socialCreditCodeOrOrganizationalCode");
        sba.append(",原注册证号,注册证号备注,分包装批准文号,公司名称（中文）,公司名称（英文）,地址（中文）,地址（英文）,国家/地区（中文）,国家/地区（英文）,产品名称（中文）,产品名称（英文）,商品名（中文）,商品名（英文）,"
                + "剂型（中文）,规格（中文）,包装规格（中文）,生产厂商（中文）,生产厂商（英文）,厂商地址（中文）,厂商地址（英文）,厂商国家/地区（中文）,厂商国家/地区（英文）,"
                + "有效期截止日,分包装企业名称,分包装企业地址,分包装文号批准日期,分包装文号有效期截止日");
        sbb.append(",originRegisterNumber,registerNumberNote,subpackageApprovalNumber,companyNameCn,companyNameEn,addressCn,addressEn,countryOrRegionCn,countryOrRegionEn,productNameCn,productNameEn,goodsNameCn,goodsNameEn,"
                + "dosageFormCn,specificationCn,packageSpecification,manufacturerCn,manufacturerEn,manufacturerAddressCn,manufacturerAddressEn,manufacturerCountryCn,manufacturerCountryEn,"
                + "expiryDate,subpackageEnterpriseName,subpackageEnterpriseAddress,subpackageApprivalNumberApprivalDate,subpackageApprivalNumberExpiryDate");
        sba.append(",使用备案号,备案状态,中药提取物信息,批准文号备注");
        sbb.append(",useRecordNumber,recordStatus,chineseMedicineExtractsInformation,approvalNumberRemark");
        sba.append(",英文名称/拉丁名称,药品分类,药品标准,备案内容");
        sbb.append(",englishNameOrLatinName,drugClassification,drugStandard,recordContent");
        sba.append(",药品本位码,药品本位码备注");
        sbb.append(",drugStandardCode,drugStandardCodeRemark");
        sba.append(",传统中药制剂名称,配制工艺路线,医疗机构注册地址");
        sbb.append(",traditionalChineseMedicinePreparationsName,processingRoute,medicalInstitutionsRegistrationAddress");
        sba.append(",产品注册证号");
        sbb.append(",productRegisterNumber");
        sba.append(",备案人名称(中文),备案人名称(原文),备案人名称(英文),代理人,代理人注册地址,产品名称中文,产品名称原文,产品名称英文,结构特征,分类编码,备案办理人,首次备案/备案变更/取消备案日期,状态,备案变更次数,备案类型,第一类医疗器械生产备案号,是否已列入生产备案凭证的产品列表");
        sbb.append(",recordHolderNameCn,recordHolderName,recordHolderNameEn,agent,agentRegistrationAddress,productNameCn,productName,productNameEn,structuralCharacteristics,classificationCode,filingAgent,firstOrChangeOrCancelFilingDate,status,filingChangesCount,filingType,firstTypeMedicalDevicesProductionFilingNumber,productionFilingProductsDocuments");
        sba.append(",内部受理号,统一审批编码,受理日期,企业注册地址,生产场所地址,注册证到期日期,产品储存条件及有效期,申请类型,产品分类,变更日期,变更内容,变更备注");
        sbb.append(",internalAcceptanceNumber,uniformApprovalCode,acceptanceDate,enterpriseRegistrationAddress,productionSiteAddress,registrationCertificateMaturityDate,productStorageConditionsValidityDate,applicationType,productClassification,changeDate,changeContent,changeNote");
        sba.append(",网站IP地址,服务器存放地址,非经营性互联网信息服务备案编号");
        sbb.append(",websiteIPAddress,serverStorageAddress,non-commercialInternetInformationServiceRecordNumber");
        sba.append(",最近一次备案日期,产品登记表,日常监督监管人员");
        sbb.append(",lastFilingDate,productRegistrationForm,dailyDupervisorsSupervisors");
        sba.append(",注册地址区县,注册街道,许可名称,发证机构,证书状态,GSP认证期至");
        sbb.append(",registeredAddressDistrict,registeredStreet,licenseName,licensingInstitution,certificateStatus,GSPCertificationPeriod");
        sba.append(",企业场所所在区,许可证(备案凭证)名称,许可证(备案凭证)编号,经营模式,监管分级,监管分类");
        sbb.append(",enterpriseArea,licenseName,licenseNumber,operationMode,supervisionGrading,supervisionClassification");
        sba.append(",产品注册文号,药包材产品名称,产品规格");
        sbb.append(",poductRegistrationNumber,productNamePharmaceuticalPackage,productSpecification");
        sba.append(",受理号,批件号,原始编号（原批准文号）,标准");
        sbb.append(",acceptanceNumber,batchNumber,originalNumber,standard");
        sba.append(",管辖区域,证书有效期至");
        sbb.append(",jurisdictionArea,certificateValidityUntil");
        sba.append(",电信业务经营许可证ICP备案号");
        sbb.append(",telecomBusinessLicenseICPRecordNumber");
        sba.append(",辖区,变更及注销记录");
        sbb.append(",jurisdiction, alterationCancellationRecords");
        sba.append(",区县,最近变更时间,生产范围(副本),质量受权人");
        sbb.append(",district,recentChangeTime,ProductionScope,qualityAuthorizedPerson");
        sba.append(",所属区县,联系电话");
        sbb.append(",district,telephone");
        sba.append(",发证/备案机关,发证/备案时间");
        sbb.append(",issuingFilingOrgan,issuingFilingTime");
        sba.append(",制剂配制单位名称,文号有效至");
        sbb.append(",namePreparationUnit,validityPeriodDate");
        sba.append(",药品生产企业,药品通用名称,文号有效期");
        sbb.append(",pharmaceuticalManufacturingEnterprises,genericNamesDrugs,validityPeriodNumbers");
        sba.append(",监督举报电话");
        sbb.append(",supervisoryReportingTelephone");
        sba.append(",产品性能结构及组成,产品适用范围,生产场所");
        sbb.append(",productPerformanceStructureComposition,productScopeApplication,productionSite");
        sba.append(",许可证有效起始日期,证有效截止日期,备案凭证编号");
        sbb.append(",licenseValidStartDate,certificateValidDeadline,recordDocumentsNumber");
        sba.append(",生产企业,编号,生产批号,企业许可证,合同编号,出口企业名称,出口企业地址,销往国家,生产数量,企业联系人");
        sbb.append(",manufacturingEnterprise,number,batchNumber,license,contractNumber,exportEnterpriseName,exportEnterpriseAddress,countrySold,productionQuantity,enterpriseContactPerson");
        sba.append(",证书号,认证范围,管理类别,注销机关,质量管理人");
        sbb.append(",certNo,certificationArea,managerType,cancellationAuthority,qualityAdministrator");
        sba.append(",批准机构,批准机关");
        sbb.append(",approvingBody,approvalAuthority ");
        sba.append(",有效截止日期,产品类型,审批机关");
        sbb.append(",effectiveDeadline,productType,approvalAuthority");
        sba.append(",委托企业,卫生许可证,委托企业地址,委托企业卫生许可证号,当日序号");
        sbb.append(",entrustedEnterprise,healthLicense,entrustedEnterpriseAddress,entrustedEnterpriseHealthLicenseNumber,dateSequenceNumber");
        sba.append(",附件");
        sbb.append(",enclosure");
        sba.append(",企业地址,许可生效期,许可截止期");
        sbb.append(",enterpriseAddress,licensingEffectivePeriod,licensingDeadline");
        sba.append(",地市,新许可证号,组织机构代码,企业类型");
        sbb.append(",city,newLicenseNumber,organizationalCode,enterpriseType");
        sba.append(",机构名称,所在地市,研究范围");
        sbb.append(",organizationName,locationCity,studyScope");
        sba.append(",商品名,英文名称,有效日期,原批准文号,执行标准");
        sbb.append(",commodityName,englishName,validDate,originalApprovalNumber,implementationStandard");
        sba.append(",通用名称,生产厂商,登记证书号,规格,药品批准文号");
        sbb.append(",generalName,manufacturer,registrationCertificateNumber,specification,drugApprovalNumber");
        sba.append(",证书有效期");
        sbb.append(",certValidityPeriod");
        sba.append(",社会信用代码(身份证号码),许可明细");
        sbb.append(",socialCreditCode,licenseDetail");
        sba.append(",社会信用代码（身份证号码）,法定代表人（负责人）");
        sbb.append(",socialCreditCode,legalPerson");
        sba.append(",企业名称（名称）");
        sbb.append(",enterpriseName");
        sba.append(",食品类型,截止日期,仓储地址");
        sbb.append(",foodType,deadline,storageAddress");
        sba.append(",社会信用代码（身份证号）,产品品种名细,有效状态");
        sbb.append(",socialCreditCode,productVarietiesDetail,effectiveStatus");
        sba.append(",发证日期证日期,许可证截止日期");
        sbb.append(",issuanceDate,licenceDeadline");
        sba.append(",法定代表负责人,监管机构,监管人员");
        sbb.append(",principalLegalRepresentative,supervisoryAuthority,supervisor");
        sba.append(",生产厂家,生产厂家所在地区,产品标准号,生产许可证编号,产地,食品添加剂,食用方法,贮存条件,保质期,配料");
        sbb.append(",manufacturers,areasManufacturersLocated,productStandardNumber,productionLicenseNumber,origin,foodAdditives,eatingMethods,storageConditions,shelfLife,ingredients");
        sba.append(",许可类别,许可证类型,许可类别说明,食品安全管理员,负责人名");
        sbb.append(",licenseType,licenseCertType,licenseTypeDesc,foodSafetyAdmin,personChargeName");
        sba.append(",负 责 人");
        sbb.append(",personCharge");
        sba.append(",产品类别,*发证机关");
        sbb.append(",productType,licensingAuthority");
        sba.append(",品名,申报单位,配制单位");
        sbb.append(",name,declarationUnit,preparationUnit");
        sba.append(",地区");
        sbb.append(",region");
        sba.append(",持有者名称,保健功能,标志性成分名称及含量,主要原料,再注册情况");
        sbb.append(",holderName,healthCareFunction,markerIngredientsName,mainRawMaterials,reregistrationStatus");
        sba.append(",省份,明细,有效期,说明");
        sbb.append(",provinces,details,validityPeriod,description");
        sba.append(",注销原因,注销日期");
        sbb.append(",cancellationReasons,cancellationDate");
        sba.append(",证书编号,检验方式,有限期至,发证单位");
        sbb.append(",certificateNumber,inspectionMethod,limitedPeriod,licensingUnit");
        sba.append(",法定代表人或主要负责人,网站ip地址,电信业务经营许可证或者非经营性互联网信息服务备案编号,医疗器械网络交易服务第三方平台备案凭证编号");
        sbb.append(",legalPerson,ip,internetRecodeNo,servicePlatformRecodeNo");
        sba.append(",发证/备案日期");
        sbb.append(",recordDate");

        sba.append(",法人,邮政编码,变更项目,变更后内容,变更时间");
        sbb.append(",legalPerson,postalCode,changeItem,changeContent,changeTime");
        sba.append(",企业地市,审批年度");
        sbb.append(",enterprisePrefecture,approvalYear");
        sba.append(",注册证号,生产单位");
        sbb.append(",registrationCertificateNumber,productionUnit");
        sba.append(",许可期限,备 注");
        sbb.append(",licensePeriod,remark");
        sba.append(",注册地址,配制地址,法 定 代表人,制 剂 负责人,医疗机构 类 别,配制范围,许可证号");
        sbb.append(",registrationAddress,preparingAddress,legalRepresentative,preparatoryOfficer,medicalInstitutionCategory,preparingScope,licenseNumber");
        sba.append(",有效期自");
        sbb.append(",validityPeriodStart");
        sba.append(",注册号,批准日期,批准有效期,发证机关,产品中文名称,产品品种编码,规格型号,注册商名称,注册厂家,注册商区域,产品结构及组成,产品主要功能,产品标准编号");
        sbb.append(",registrationNumber,approvalDate,approvalValidityPeriod,licensingAuthority,productChineseName,productVarietyCode,specificationModel,registererName,registeredManufacturer,registererArea,productStructurecomposition,productMainFunctions,productStandardNumber");
        sba.append(",企业地区名称,生产负责人");
        sbb.append(",enterpriseAreaName,productLeader");
        sba.append(",企业所在地区名称,生产许可证号,许可机关");
        sbb.append(",enterpriseAreaName,productionLicenseNumber,licensingAuthority");
        sba.append(",配置地址,配制范围,业务负责人");
        sbb.append(",configurationAddress,configurationScope,businessLeader");
        sba.append(",GSP证号,GSP发证日期,GSP有效期至");
        sbb.append(",GSPCertificateNo,GSPIssuingCertificateDate,GSPCertificateExpirationDate");
        sba.append(",地区名称,许可证发证日期,许可证有效日期止,生产地址及生产范围,相关数据库查询,注");
        sbb.append(",areaName,licenseIssuingDate,licenseValidityDate,productionAddressScopeProduction,relevantDatabaseQuery,note");

        sba.append(",注册证编号,注册人名称,注册人住所,代理人名称,代理人住所,型号、规格/包装规格,结构及组成/主要组成成分,适用范围/预期用途,产品储存条件及有效期(体外诊断试剂适用),其他内容,审批部门,批证日期,变更信息");
        sbb.append(",registrationCertificateNumber,registrarName,registrarResidence,agentName,agentResidence,model,specificationOrPackagingSpecification,applicableScopeOrExpectedUse,productStorageConditionsValidityPeriod,otherContents,approvalDepartment,approvalDate,changeInformation");
        sba.append(",备案人名称,产品类名称,包装规格,产品有效期,主要组成成分");
        sbb.append(",nameFiler,nameProductCategory,packingSpecification,validityPeriodProduct,mainComponents");
        sba.append(",备案人组织机构代码,备案人注册地址,型号/规格,产品描述,预期用途,变更情况,检查监督情况");
        sbb.append(",organizationalCodeRecorder,registrationAddress,modelSpecification,productDescription,expectedUse,change,inspectionSupervisionRecorder");
        sba.append(",变更记录,许可证状态");
        sbb.append(",changeRecord,licenseStatus");
        sba.append(",生产范围,备案单位,产品信息");
        sbb.append(",productionScope,recordingUnit,productInformation");
        sba.append(",备案编号,备案部门,备案日期");
        sbb.append(",recordNumber,recordDepartment,recordDate");
        sba.append(",库房地址,发证部门,有效期限");
        sbb.append(",warehouseAddress,issuingDepartment,validityPeriod");
        sba.append(",企业名称,标准名称,标准编号,标准备案编号,备案时间,备案有效期,标准全文和修订单");
        sbb.append(",enterpriseName,standardName,standardNumber,standardFilingNumber,filingTime,filingValidityPeriod,standardFullTextRevisionForm");
        sba.append(",单位名称,法定代表人或业主,证件状态");
        sbb.append(",unitName,legalRepresentativeOwner,certificateStatus");
        sba.append(",生产地址和生产范围");
        sbb.append(",productionAddressAndScope");

        sba.append(",GSP证书,GSP证书发证日期,GSP证书有效期至");
        sbb.append(",GSPCertificate,issuingCertificateDate,certificateExpirationDate");
        sba.append(",制剂室负责人");
        sbb.append(",headPharmaceuticalDepartment");
        sba.append(",配制地址和配制范围");
        sbb.append(",configurationAddressAndscope ");
        sba.append(",仓库地址,经营方式,经营范围,GSP证书编号,GSP认证日期,GSP证书有效期");
        sbb.append(",warehouseAddress,modeOperation,businessScope,GSPCertificateNumber,GSPCertificationDate,GSPCertificateValidityPeriod");

        sba.append(",批准文号,制剂名称,配制单位所属辖区,制剂类别,剂型,规格,配制单位名称");
        sbb.append(",approvalNumber,preparationName,jurisdictionPreparationUnit,categoryPreparation,dosageForm,specification,preparationUnitName");

        sba.append(",主要负责人,医疗器械质量安全管理人,办公场所,网站名称,网站域名,网络客户端应用程序名,现场检查结果");
        sbb.append(",responsiblePerson,medicalDeviceQualitySafetyAdministrator,officeSpace,websiteName,websiteDomainName,networkClientApplicationName,onSiteInspectionResults");

        sba.append(",序号,备案号,备案药品生产企业名称,调剂毒性中药饮品片品种名称,备案有效期");
        sbb.append(",serialNumber,recordNumber,recordDrugManufacturerName,varietyDispensingToxicChineseMedicineDrinkName,filingTime");
        sba.append(",许可项目");
        sbb.append(",licensingProject");
        sba.append(",经营场所或生产场所,医疗器械网络销售类型,医疗器械生产（经营）许可证或备案凭证编号,医疗器械网络交易服务第三方平台名称");
        sbb.append(",businessPlaceOrProductionPlace,networkSalesMedicalDevicesType,numberLicenseOrRecordCertificateProductionMedicalDevices,thirdPartyPlatformName");
        sba.append(",域名,IP地址,电信业务经营许可证号,法定代表人或负责人");
        sbb.append(",domain,IPAddress,telecommunicationsBusinessLicenseNumber,legalRepresentativePersonCharge");
        sba.append(",备案人,备案人地址,备案结论");
        sbb.append(",filer,addressFiler,conclusionFiling");
        sba.append(",备案机关,经营负责人姓名,食品安全管理人姓名,经营范围和方式,经营地址");
        sbb.append(",recordFilingOrgans,namesBusinessLeaders,foodSafetyAdministratorName,businessScopeMethods,businessAddresses");
        String a = sba.toString();
        String b = sbb.toString();

        String[] names = a.split(",");
        String[] nameEns = b.split(",");
        for (int i = 0; i < names.length; i++) {
            map.put(names[i], nameEns[i]);
        }
        System.out.println("\n");

        if (true) {
            String[] names2 = cn2.split(",");
            System.out.println(cn2);
            List<String> unTranslate = new ArrayList<>();
            for (int i = 0; i < names2.length; i++) {
                String nameEn = map.get(names2[i]);
                if (nameEn == null || nameEn.length() == 0) {
                    nameEn = "xxxxxxxxxxxxxxxxxx" + names2[i];
                    unTranslate.add(names2[i]);
                }
                if ((i + 1) == names2.length) {
                    System.out.print(nameEn);
                } else {
                    System.out.print(nameEn + ",");

                }
            }
            System.out.println();
            for (String string : unTranslate) {
                System.out.print(string + ",");
            }

            System.out.println("\n");

            System.out.println();
            System.out.println(names.length);
            System.out.println(nameEns.length);
            System.out.println(map);
//            for (int i = 0; i < names.length; i++) {
//                System.out.print(names[i] + "=" + nameEns[i]);
//            }

        }

    }

    public static void main(String[] args) {

    }

}
