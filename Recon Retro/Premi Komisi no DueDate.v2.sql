
create table #temp
(
	
	PolicyNo varchar(10),
	PolicyTransactionNo smallint,
	IssueDate date,
	EffectiveDate date,
	PPWDays smallint,
	DueDate date,
	Premium numeric(20,5) default(0.0),
	Discount numeric(20,5) default(0.0),
	Commission numeric(20,5) default(0.0)

)

insert into #temp
(PolicyNo, PolicyTransactionNo, IssueDate, EffectiveDate, PPWDays, DueDate, Premium, Discount, Commission)
select p.PolicyNo, p.TransactionNo, dd.CalendarDate as IssueDate, pd.EffectiveDate, x.PPWDays, rpid.DueDate, p.OriginalCurrencyPremium as Premium, p.OriginalCurrencyDiscount as Discount, p.OriginalCurrencyCommission as Commission
from dbo.Productions p
left join dbo.Risks r
	left join dbo.PolicyDimension pd
	on (r.PolicySID = r.PolicySID)
	left join dbo.DateDimension dd
	on (r.IssueDateSID = dd.DateSID)
on (p.PolicyNo = pd.PolicyNo and p.TransactionNo = pd.TransactionNo)
left join
	(select xpt.PolicyNo, xpt.PolicyTransactionNo, xpr.PPW as PPWDays
	from [TMI.Actuary].dbo.PolicyTransactions xpt
	inner join [TMI.Raw].dbo.PolicyDimension xpd
	on (xpt.PolicyNo = xpd.PolicyNo and xpt.PolicyTransactionNo = xpd.TransactionNo)
	inner join [TMI.Raw].dbo.PolicyClauseDimension xpcd
	on (xpd.PolicySID = xpcd.PolicySID)
	inner join [TMI.Actuary].dbo.PPWRetro xpr
	on (concat(xpt.ContractTypeCode, xpcd.ClauseCode) = xpr.ContractClause))x
on (p.PolicyNo = x.PolicyNo and p.TransactionNo = x.PolicyTransactionNo)
left join [TMI.Raw].dbo.PolicyInstallmentDimension rpid
on (p.PolicyNo = rpid.PolicyNo and p.TransactionNo = rpid.TransactionNo)
where dd.CalendarDate between '2023-01-01' and '2023-12-31'
group by p.PolicyNo, p.TransactionNo, dd.CalendarDate, pd.EffectiveDate, x.PPWDays, rpid.DueDate, p.OriginalCurrencyPremium, p.OriginalCurrencyDiscount, p.OriginalCurrencyCommission
order by p.PolicyNo, p.TransactionNo, dd.CalendarDate, pd.EffectiveDate, x.PPWDays, rpid.DueDate, p.OriginalCurrencyPremium, p.OriginalCurrencyDiscount, p.OriginalCurrencyCommission

select xpt.PolicyNo, xpt.PolicyTransactionNo, xpr.PPW as PPWDays
	from [TMI.Actuary].dbo.PolicyTransactions xpt
	inner join [TMI.Raw].dbo.PolicyDimension xpd
	on (xpt.PolicyNo = xpd.PolicyNo and xpt.PolicyTransactionNo = xpd.TransactionNo)
	inner join [TMI.Raw].dbo.PolicyClauseDimension xpcd
	on (xpd.PolicySID = xpcd.PolicySID)
	inner join [TMI.Actuary].dbo.PPWRetro xpr
	on (concat(xpt.ContractTypeCode, xpcd.ClauseCode) = xpr.ContractClause)


delete from #temp
where DueDate is not null

update t
set t.PPWDays = isnull(t.PPWDays,0), t.Commission = isnull(t.Commission, 0), t.Premium = isnull(t.Premium,0)
from #temp t

update t
set t.DueDate =
	case 
		when t.EffectiveDate > t.IssueDate 
		then dateadd(day, t.PPWDays, t.EffectiveDate)
		else dateadd(day, t.PPWDays, t.IssueDate)
	end
from #temp t

select t.PolicyNo, t.PolicyTransactionNo, t.IssueDate, t.DueDate, t.Premium, t.Commission
from #temp t
order by t.PolicyNo, t.PolicyTransactionNo, t.IssueDate, t.DueDate

drop table #temp