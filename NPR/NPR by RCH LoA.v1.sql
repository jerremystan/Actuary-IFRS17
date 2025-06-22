

--Tabel NXLTFO, semua Long Term
select 'TFO' as ReinsuranceTypeCode
		, it.ReservingClassCode
		, 0 as IsShortTerm
		, tota.InsurerCode
		, year(ic.IssueDate) as Cohort
		, sum(isnull(tota.NetRetainedPremium,0) + isnull(tota.Discount,0)) as RWP
from dbo.InwardTransactions it
	inner join dbo.TreatyOutwardTransactionAmount tota on (it.PolicyNo = tota.PolicyNo and it.PolicyTransactionNo = tota.PolicyTransactionNo)
	inner join dbo.InsuranceContractRecognitions ic on (it.PolicyNo = ic.PolicyNo and it.RenewalNo = ic.RenewalNo)
where year(ic.IssueDate) = '2022'
group by it.ReservingClassCode
		, tota.InsurerCode
		, year(ic.IssueDate)

union all

--Tabel XL, 
select 'XL' as ReinsuranceTypeCode
		, npt.ReservingClassCode
		, npc.IsShortTerm
		, tota.InsurerCode
		, year(npcr.IssueDate) as Cohort
		,  sum(isnull(tota.NetRetainedPremium,0) + isnull(tota.Discount,0)) as RWP
from dbo.NonProportionalTransactions npt
	inner join dbo.TreatyOutwardTransactionAmount tota on (npt.PolicyNo = tota.PolicyNo and npt.PolicyTransactionNo = npt.PolicyTransactionNo)
	inner join dbo.NonProportionalContractRecognitions npcr 
		inner join dbo.NonProportionalContracts npc on (npcr.PolicyNo = npc.PolicyNo and npcr.RenewalNo = npc.RenewalNo)
		on (npt.PolicyNo = npcr.PolicyNo and npt.RenewalNo = npcr.RenewalNo)
where year(npcr.IssueDate) = '2022'
group by npt.ReservingClassCode
		, npc.IsShortTerm
		, tota.InsurerCode
		, year(npcr.IssueDate)

go


--Tabel NXLF, Faculcative
select 'F' as ReinsuranceTypeCode
		, it.ReservingClassCode
		, it.IsShortTerm
		, fota.PolicyNo
		, it.RenewalNo
		--, fota.RiskNo
		, fota.InsurerCode
		, year(icr.IssueDate) as Cohort
		,  sum(isnull(fota.NetRetainedPremium,0) + isnull(fota.Discount,0)) as RWP
from dbo.InwardTransactions it
	inner join dbo.FacultativeOutwardTransactionAmount fota on (it.PolicyNo = fota.PolicyNo and it.PolicyTransactionNo = fota.PolicyTransactionNo)
	inner join dbo.InsuranceContractRecognitions icr 
		inner join dbo.InsuranceContracts ic on (icr.PolicyNo = ic.PolicyNo and icr.RenewalNo = ic.RenewalNo)
		on (it.PolicyNo = icr.PolicyNo and it.RenewalNo = icr.RenewalNo)
where year(icr.IssueDate) = '2022'
group by it.ReservingClassCode
		, it.IsShortTerm
		, fota.PolicyNo
		, it.RenewalNo
		--, fota.RiskNo
		, fota.InsurerCode
		, year(icr.IssueDate)


/* check treaty code NXLTFO
select distinct tota.InsurerCode
from dbo.InwardTransactions it
	inner join dbo.InsuranceContractRecognitions ic on (it.PolicyNo = ic.PolicyNo and it.RenewalNo = ic.RenewalNo)
	inner join dbo.TreatyOutwardTransactionAmount tota on (it.PolicyNo = tota.PolicyNo and it.PolicyTransactionNo = tota.PolicyTransactionNo)
	
where year(ic.IssueDate) = '2023'
order by tota.InsurerCode
*/
