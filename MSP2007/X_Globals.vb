Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Module X_Globals

	'// Enumerations 

	Public Enum E_FYSTARTDATE
		FYSD_JANUARY = 1
		FYSD_FEBRUARY = 2
		FYSD_MARCH = 3
		FYSD_APRIL = 4
		FYSD_MAY = 5
		FYSD_JUNE = 6
		FYSD_JULY = 7
		FYSD_AUGUST = 8
		FYSD_SEPTEMBER = 9
		FYSD_OCTOBER = 10
		FYSD_NOVEMBER = 11
		FYSD_DECEMBER = 12
	End Enum

	Public Enum E_CURRENCYSYMBOLPOSITION
		CSP_BEFORE = 0
		CSP_AFTER = 1
		CSP_BEFORE_WITH_SPACE = 2
		CSP_AFTER_WITH_SPACE = 3
	End Enum

	Public Enum E_DEFAULTTASKTYPE
		DTT_FIXED_UNITS = 0
		DTT_FIXED_DURATION = 1
		DTT_FIXED_WORK = 2
	End Enum

	Public Enum E_DEFAULTFIXEDCOSTACCRUAL
		DFCA_START = 1
		DFCA_PRORATED = 2
		DFCA_END = 3
	End Enum

	Public Enum E_DURATIONFORMAT
		DF_M = 3
		DF_EM = 4
		DF_H = 5
		DF_EH = 6
		DF_D = 7
		DF_ED = 8
		DF_W = 9
		DF_EW = 10
		DF_MO = 11
		DF_EMO = 12
		DF__PERCENT = 19
		DF_E_PERCENT = 20
		DF_NULL = 21
		DF_M_QUESTIONMRK = 35
		DF_EM_QUESTIONMRK = 36
		DF_H_QUESTIONMRK = 37
		DF_EH_QUESTIONMRK = 38
		DF_D_QUESTIONMRK = 39
		DF_ED_QUESTIONMRK = 40
		DF_W_QUESTIONMRK = 41
		DF_EW_QUESTIONMRK = 42
		DF_MO_QUESTIONMRK = 43
		DF_EMO_QUESTIONMRK = 44
		DF__PERCENT_QUESTIONMRK = 51
		DF_E_PERCENT_QUESTIONMRK = 52
		DF_NULL_A = 53
	End Enum

	Public Enum E_WORKFORMAT
		WF_M = 1
		WF_H = 2
		WF_D = 3
		WF_W = 4
		WF_MO = 5
	End Enum

	Public Enum E_EARNEDVALUEMETHOD
		EVM_PERCENT_COMPLETE = 0
		EVM_PHYSICAL_PERCENT_COMPLETE = 1
	End Enum

	Public Enum E_WEEKSTARTDAY
		WSD_SUNDAY = 0
		WSD_MONDAY = 1
		WSD_TUESDAY = 2
		WSD_WEDNESDAY = 3
		WSD_THURSDAY = 4
		WSD_FRIDAY = 5
		WSD_SATURDAY = 6
	End Enum

	Public Enum E_BASELINEFOREARNEDVALUE
		BFEV_BASELINE = 0
		BFEV_BASELINE_1 = 1
		BFEV_BASELINE_2 = 2
		BFEV_BASELINE_3 = 3
		BFEV_BASELINE_4 = 4
		BFEV_BASELINE_5 = 5
		BFEV_BASELINE_6 = 6
		BFEV_BASELINE_7 = 7
		BFEV_BASELINE_8 = 8
		BFEV_BASELINE_9 = 9
		BFEV_BASELINE_10 = 10
	End Enum

	Public Enum E_NEWTASKSTARTDATE
		NTSD_PROJECT_START_DATE = 0
		NTSD_CURRENT_DATE = 1
	End Enum

	Public Enum E_DEFAULTTASKEVMETHOD
		DTEVM_PERCENT_COMPLETE = 0
		DTEVM_PHYSICAL_PERCENT_COMPLETE = 1
	End Enum

	Public Enum E_TYPE
		T_DATE = 4
		T_DURATION = 6
		T_COST = 9
		T_NUMBER = 15
		T_FLAG = 17
		T_TEXT = 21
		T_FINISH_DATE = 27
	End Enum

	Public Enum E_TYPE_1
		T_1_NUMBERS = 0
		T_1_UPPER_CASE_LETTERS = 1
		T_1_LOWER_CASE_LETTERS = 2
		T_1_CHARACTERS = 3
	End Enum

	Public Enum E_TYPE_2
		T_2_NUMBERS = 0
		T_2_UPPERCASE_LETTERS = 1
		T_2_LOWERCASE_LETTERS = 2
		T_2_CHARACTERS = 3
	End Enum

	Public Enum E_CFTYPE
		CFT_COST = 0
		CFT_DATE = 1
		CFT_DURATION = 2
		CFT_FINISH = 3
		CFT_FLAG = 4
		CFT_NUMBER = 5
		CFT_START = 6
		CFT_TEXT = 7
	End Enum

	Public Enum E_ELEMTYPE
		ET_TASK = 20
		ET_RESOURCE = 21
		ET_CALENDAR = 22
		ET_ASSIGNMENT = 23
	End Enum

	Public Enum E_ROLLUPTYPE
		RT_MAXIMUM_OR_FOR_FLAG_FIELDS = 0
		RT_MINIMUM__AND__FOR_FLAG_FIELDS = 1
		RT_COUNT_ALL = 2
		RT_SUM = 3
		RT_AVERAGE = 4
		RT_AVERAGE_FIRST_SUBLEVEL = 5
		RT_COUNT_FIRST_SUBLEVEL = 6
		RT_COUNT_NONSUMMARIES = 7
	End Enum

	Public Enum E_CALCULATIONTYPE
		CT_NONE = 0
		CT_ROLLUP = 1
		CT_CALCULATION = 2
	End Enum

	Public Enum E_VALUELISTSORTORDER
		VSO_DESCENDING = 0
		VSO_ASCENDING = 1
	End Enum

	Public Enum E_DAYTYPE
		DT_EXCEPTION = 0
		DT_SUNDAY = 1
		DT_MONDAY = 2
		DT_TUESDAY = 3
		DT_WEDNESDAY = 4
		DT_THURSDAY = 5
		DT_FRIDAY = 6
		DT_SATURDAY = 7
	End Enum

	Public Enum E_TYPE_3
		T_3_DAILY = 1
		T_3_YEARLY_BY_DAY_OF_THE_MONTH = 2
		T_3_YEARLY_BY_POSITION = 3
		T_3_MONTHLY_BY_DAY_OF_THE_MONTH = 4
		T_3_MONTHLY_BY_POSITION = 5
		T_3_WEEKLY = 6
		T_3_BY_DAY_COUNT = 7
		T_3_BY_WEEKDAY_COUNT = 8
		T_3_NO_EXCEPTION_TYPE = 9
	End Enum

	Public Enum E_MONTHITEM
		MI_DAY = 0
		MI_WEEKDAY = 1
		MI_WEEKENDDAY = 2
		MI_SUNDAY = 3
		MI_MONDAY = 4
		MI_TUESDAY = 5
		MI_WEDNESDAY = 6
		MI_THURSDAY = 7
		MI_FRIDAY = 8
		MI_SATURDAY = 9
	End Enum

	Public Enum E_MONTHPOSITION
		MP_FIRST_POSITION = 0
		MP_SECOND_POSITION = 1
		MP_THIRD_POSITION = 2
		MP_FOURTH_POSITION = 3
		MP_LAST_POSITION = 4
	End Enum

	Public Enum E_MONTH
		M_JANUARY = 0
		M_FEBRUARY = 1
		M_MARCH = 2
		M_APRIL = 3
		M_MAY = 4
		M_JUNE = 5
		M_JULY = 6
		M_AUGUST = 7
		M_SEPTEMBER = 8
		M_OCTOBER = 9
		M_NOVEMBER = 10
		M_DECEMBER = 11
	End Enum

	Public Enum E_TYPE_4
		T_4_FIXED_UNITS = 0
		T_4_FIXED_DURATION = 1
		T_4_FIXED_WORK = 2
	End Enum

	Public Enum E_FIXEDCOSTACCRUAL
		FCA_START = 1
		FCA_PRORATED = 2
		FCA_END = 3
	End Enum

	Public Enum E_CONSTRAINTTYPE
		CT_AS_SOON_AS_POSSIBLE = 0
		CT_AS_LATE_AS_POSSIBLE = 1
		CT_MUST_START_ON = 2
		CT_MUST_FINISH_ON = 3
		CT_START_NO_EARLIER_THAN = 4
		CT_START_NO_LATER_THAN = 5
		CT_FINISH_NO_EARLIER_THAN = 6
		CT_FINISH_NO_LATER_THAN = 7
	End Enum

	Public Enum E_LEVELINGDELAYFORMAT
		LDF_M = 3
		LDF_EM = 4
		LDF_H = 5
		LDF_EH = 6
		LDF_D = 7
		LDF_ED = 8
		LDF_W = 9
		LDF_EW = 10
		LDF_MO = 11
		LDF_EMO = 12
		LDF__PERCENT = 19
		LDF_E_PERCENT = 20
		LDF_NULL = 21
		LDF_M_QUESTIONMRK = 35
		LDF_EM_QUESTIONMRK = 36
		LDF_H_QUESTIONMRK = 37
		LDF_EH_QUESTIONMRK = 38
		LDF_D_QUESTIONMRK = 39
		LDF_ED_QUESTIONMRK = 40
		LDF_W_QUESTIONMRK = 41
		LDF_EW_QUESTIONMRK = 42
		LDF_MO_QUESTIONMRK = 43
		LDF_EMO_QUESTIONMRK = 44
		LDF__PERCENT_QUESTIONMRK = 51
		LDF_E_PERCENT_QUESTIONMRK = 52
		LDF_NULL_A = 53
	End Enum

	Public Enum E_COMMITMENTTYPE
		CT_THE_TASK_HAS_NO_DELIVERABLE_OR_DEPENDENCY_ON_A_DELIVERABLE = 0
		CT_THE_TASK_HAS_AN_ASSOCIATED_DELIVERABLE = 1
		CT_THE_TASK_HAS_A_DEPENDENCY_ON_AN_ASSOCIATED_DELIVERABLE = 2
	End Enum

	Public Enum E_TYPE_5
		T_5_FF = 0
		T_5_FS = 1
		T_5_SF = 2
		T_5_SS = 3
	End Enum

	Public Enum E_LAGFORMAT
		LF_M = 3
		LF_EM = 4
		LF_H = 5
		LF_EH = 6
		LF_D = 7
		LF_ED = 8
		LF_W = 9
		LF_EW = 10
		LF_MO = 11
		LF_EMO = 12
		LF__PERCENT = 19
		LF_E_PERCENT = 20
		LF_M_QUESTIONMRK = 35
		LF_EM_QUESTIONMRK = 36
		LF_H_QUESTIONMRK = 37
		LF_EH_QUESTIONMRK = 38
		LF_D_QUESTIONMRK = 39
		LF_ED_QUESTIONMRK = 40
		LF_W_QUESTIONMRK = 41
		LF_EW_QUESTIONMRK = 42
		LF_MO_QUESTIONMRK = 43
		LF_EMO_QUESTIONMRK = 44
		LF__PERCENT_QUESTIONMRK = 51
		LF_E_PERCENT_QUESTIONMRK = 52
	End Enum

	Public Enum E_TYPE_6
		T_6_MATERIAL = 0
		T_6_WORK = 1
	End Enum

	Public Enum E_WORKGROUP
		WG_DEFAULT = 0
		WG_NONE = 1
		WG_EMAIL = 2
		WG_WEB = 3
	End Enum

	Public Enum E_ACCRUEAT
		AA_START = 1
		AA_END = 2
		AA_PRORATED = 3
		AA_INVALID = 4
	End Enum

	Public Enum E_STANDARDRATEFORMAT
		SRF_M = 1
		SRF_H = 2
		SRF_D = 3
		SRF_W = 4
		SRF_MO = 5
		SRF_Y = 7
		SRF_MATERIAL_RESOURCE_RATE_OR_BLANK_SYMBOL_SPECIFIED = 8
	End Enum

	Public Enum E_OVERTIMERATEFORMAT
		ORF_M = 1
		ORF_H = 2
		ORF_D = 3
		ORF_W = 4
		ORF_MO = 5
		ORF_Y = 7
	End Enum

	Public Enum E_BOOKINGTYPE
		BT_COMMITED = 1
		BT_PROPOSED = 2
	End Enum

	Public Enum E_RATETABLE
		RT_A = 0
		RT_B = 1
		RT_C = 2
		RT_D = 3
		RT_E = 4
	End Enum

	Public Enum E_STANDARDRATEFORMAT_1
		SRF_1_M = 1
		SRF_1_H = 2
		SRF_1_D = 3
		SRF_1_W = 4
		SRF_1_MO = 5
		SRF_1_Y = 7
	End Enum

	Public Enum E_COSTRATETABLE
		CRT_COST_RATE_TABLE_0 = 0
		CRT_COST_RATE_TABLE_1 = 1
		CRT_COST_RATE_TABLE_2 = 2
		CRT_COST_RATE_TABLE_3 = 3
		CRT_COST_RATE_TABLE_4 = 4
	End Enum

	Public Enum E_WORKCONTOUR
		WC_FLAT = 0
		WC_BACK_LOADED = 1
		WC_FRONT_LOADED = 2
		WC_DOUBLE_PEAK = 3
		WC_EARLY_PEAK = 4
		WC_LATE_PEAK = 5
		WC_BELL = 6
		WC_TURTLE = 7
		WC_CONTOURED = 8
	End Enum

	Public Enum E_TYPE_7
		T_7_ASSIGNMENT_REMAINING_WORK = 1
		T_7_ASSIGNMENT_ACTUAL_WORK = 2
		T_7_ASSIGNMENT_ACTUAL_OVERTIME_WORK = 3
		T_7_ASSIGNMENT_BASELINE_WORK = 4
		T_7_ASSIGNMENT_BASELINE_COST = 5
		T_7_ASSIGNMENT_ACTUAL_COST = 6
		T_7_RESOURCE_BASELINE_WORK = 7
		T_7_RESOURCE_BASELINE_COST = 8
		T_7_TASK_BASELINE_WORK = 9
		T_7_TASK_BASELINE_COST = 10
		T_7_TASK_PERCENT_COMPLETE = 11
		T_7_ASSIGNMENT_BASELINE_1_WORK = 16
		T_7_ASSIGNMENT_BASELINE_1_COST = 17
		T_7_TASK_BASELINE_1_WORK = 18
		T_7_TASK_BASELINE_1_COST = 19
		T_7_RESOURCE_BASELINE_1_WORK = 20
		T_7_RESOURCE_BASELINE_1_COST = 21
		T_7_ASSIGNMENT_BASELINE_2_WORK = 22
		T_7_ASSIGNMENT_BASELINE_2_COST = 23
		T_7_TASK_BASELINE_2_WORK = 24
		T_7_TASK_BASELINE_2_COST = 25
		T_7_RESOURCE_BASELINE_2_WORK = 26
		T_7_RESOURCE_BASELINE_2_COST = 27
		T_7_ASSIGNMENT_BASELINE_3_WORK = 28
		T_7_ASSIGNMENT_BASELINE_3_COST = 29
		T_7_TASK_BASELINE_3_WORK = 30
		T_7_TASK_BASELINE_3_COST = 31
		T_7_RESOURCE_BASELINE_3_WORK = 32
		T_7_RESOURCE_BASELINE_3_COST = 33
		T_7_ASSIGNMENT_BASELINE_4_WORK = 34
		T_7_ASSIGNMENT_BASELINE_4_COST = 35
		T_7_TASK_BASELINE_4_WORK = 36
		T_7_TASK_BASELINE_4_COST = 37
		T_7_RESOURCE_BASELINE_4_WORK = 38
		T_7_RESOURCE_BASELINE_4_COST = 39
		T_7_ASSIGNMENT_BASELINE_5_WORK = 40
		T_7_ASSIGNMENT_BASELINE_5_COST = 41
		T_7_TASK_BASELINE_5_WORK = 42
		T_7_TASK_BASELINE_5_COST = 43
		T_7_RESOURCE_BASELINE_5_WORK = 44
		T_7_RESOURCE_BASELINE_5_COST = 45
		T_7_ASSIGNMENT_BASELINE_6_WORK = 46
		T_7_ASSIGNMENT_BASELINE_6_COST = 47
		T_7_TASK_BASELINE_6_WORK = 48
		T_7_TASK_BASELINE_6_COST = 49
		T_7_RESOURCE_BASELINE_6_WORK = 50
		T_7_RESOURCE_BASELINE_6_COST = 51
		T_7_ASSIGNMENT_BASELINE_7_WORK = 52
		T_7_ASSIGNMENT_BASELINE_7_COST = 53
		T_7_TASK_BASELINE_7_WORK = 54
		T_7_TASK_BASELINE_7_COST = 55
		T_7_RESOURCE_BASELINE_7_WORK = 56
		T_7_RESOURCE_BASELINE_7_COST = 57
		T_7_ASSIGNMENT_BASELINE_8_WORK = 58
		T_7_ASSIGNMENT_BASELINE_8_COST = 59
		T_7_TASK_BASELINE_8_WORK = 60
		T_7_TASK_BASELINE_8_COST = 61
		T_7_RESOURCE_BASELINE_8_WORK = 62
		T_7_RESOURCE_BASELINE_8_COST = 63
		T_7_ASSIGNMENT_BASELINE_9_WORK = 64
		T_7_ASSIGNMENT_BASELINE_9_COST = 65
		T_7_TASK_BASELINE_9_WORK = 66
		T_7_TASK_BASELINE_9_COST = 67
		T_7_RESOURCE_BASELINE_9_WORK = 68
		T_7_RESOURCE_BASELINE_9_COST = 69
		T_7_ASSIGNMENT_BASELINE_10_WORK = 70
		T_7_ASSIGNMENT_BASELINE_10_COST = 71
		T_7_TASK_BASELINE_10_WORK = 72
		T_7_TASK_BASELINE_10_COST = 73
		T_7_RESOURCE_BASELINE_10_WORK = 74
		T_7_RESOURCE_BASELINE_10_COST = 75
		T_7_PHYSICAL_PERCENT_COMPLETE = 76
	End Enum

	Public Enum E_UNIT
		U_M = 0
		U_H = 1
		U_D = 2
		U_W = 3
		U_MO = 5
		U_Y = 8
	End Enum


End Module
