/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

// Import
var History = AscCommon.History;


/** @enum {number} */
var c_oAscMathType = {
    //----------------------------------------------------------------------------------------------------------------------
    Symbol_pm                               : 0x00000000,
    Symbol_infinity                         : 0x00000001,
    Symbol_equals                           : 0x00000002,
    Symbol_neq                              : 0x00000003,
    Symbol_about                            : 0x00000004,
    Symbol_times                            : 0x00000005,
    Symbol_div                              : 0x00000006,
    Symbol_factorial                        : 0x00000007,
    Symbol_propto                           : 0x00000008,
    Symbol_less                             : 0x00000009,
    Symbol_ll                               : 0x0000000a,
    Symbol_greater                          : 0x0000000b,
    Symbol_gg                               : 0x0000000c,
    Symbol_leq                              : 0x0000000d,
    Symbol_geq                              : 0x0000000e,
    Symbol_mp                               : 0x0000000f,
    Symbol_cong                             : 0x00000010,
    Symbol_approx                           : 0x00000011,
    Symbol_equiv                            : 0x00000012,
    Symbol_forall                           : 0x00000013,
    Symbol_additional                       : 0x00000014,
    Symbol_partial                          : 0x00000015,
    Symbol_sqrt                             : 0x00000016,
    Symbol_cbrt                             : 0x00000017,
    Symbol_qdrt                             : 0x00000018,
    Symbol_cup                              : 0x00000019,
    Symbol_cap                              : 0x0000001a,
    Symbol_emptyset                         : 0x0000001b,
    Symbol_percent                          : 0x0000001c,
    Symbol_degree                           : 0x0000001d,
    Symbol_fahrenheit                       : 0x0000001e,
    Symbol_celsius                          : 0x0000001f,
    Symbol_inc                              : 0x00000020,
    Symbol_nabla                            : 0x00000021,
    Symbol_exists                           : 0x00000022,
    Symbol_notexists                        : 0x00000023,
    Symbol_in                               : 0x00000024,
    Symbol_ni                               : 0x00000025,
    Symbol_leftarrow                        : 0x00000026,
    Symbol_uparrow                          : 0x00000027,
    Symbol_rightarrow                       : 0x00000028,
    Symbol_downarrow                        : 0x00000029,
    Symbol_leftrightarrow                   : 0x0000002a,
    Symbol_therefore                        : 0x0000002b,
    Symbol_plus                             : 0x0000002c,
    Symbol_minus                            : 0x0000002d,
    Symbol_not                              : 0x0000002e,
    Symbol_ast                              : 0x0000002f,
    Symbol_bullet                           : 0x00000030,
    Symbol_vdots                            : 0x00000031,
    Symbol_cdots                            : 0x00000032,
    Symbol_rddots                           : 0x00000033,
    Symbol_ddots                            : 0x00000034,
    Symbol_aleph                            : 0x00000035,
    Symbol_beth                             : 0x00000036,
    Symbol_QED                              : 0x00000037,
    Symbol_alpha                            : 0x00010000,
    Symbol_beta                             : 0x00010001,
    Symbol_gamma                            : 0x00010002,
    Symbol_delta                            : 0x00010003,
    Symbol_varepsilon                       : 0x00010004,
    Symbol_epsilon                          : 0x00010005,
    Symbol_zeta                             : 0x00010006,
    Symbol_eta                              : 0x00010007,
    Symbol_theta                            : 0x00010008,
    Symbol_vartheta                         : 0x00010009,
    Symbol_iota                             : 0x0001000a,
    Symbol_kappa                            : 0x0001000b,
    Symbol_lambda                           : 0x0001000c,
    Symbol_mu                               : 0x0001000d,
    Symbol_nu                               : 0x0001000e,
    Symbol_xsi                              : 0x0001000f,
    Symbol_o                                : 0x00010010,
    Symbol_pi                               : 0x00010011,
    Symbol_varpi                            : 0x00010012,
    Symbol_rho                              : 0x00010013,
    Symbol_varrho                           : 0x00010014,
    Symbol_sigma                            : 0x00010015,
    Symbol_varsigma                         : 0x00010016,
    Symbol_tau                              : 0x00010017,
    Symbol_upsilon                          : 0x00010018,
    Symbol_varphi                           : 0x00010019,
    Symbol_phi                              : 0x0001001a,
    Symbol_chi                              : 0x0001001b,
    Symbol_psi                              : 0x0001001c,
    Symbol_omega                            : 0x0001001d,
    Symbol_Alpha                            : 0x00020000,
    Symbol_Beta                             : 0x00020001,
    Symbol_Gamma                            : 0x00020002,
    Symbol_Delta                            : 0x00020003,
    Symbol_Epsilon                          : 0x00020004,
    Symbol_Zeta                             : 0x00020005,
    Symbol_Eta                              : 0x00020006,
    Symbol_Theta                            : 0x00020007,
    Symbol_Iota                             : 0x00020008,
    Symbol_Kappa                            : 0x00020009,
    Symbol_Lambda                           : 0x0002000a,
    Symbol_Mu                               : 0x0002000b,
    Symbol_Nu                               : 0x0002000c,
    Symbol_Xsi                              : 0x0002000d,
    Symbol_O                                : 0x0002000e,
    Symbol_Pi                               : 0x0002000f,
    Symbol_Rho                              : 0x00020010,
    Symbol_Sigma                            : 0x00020011,
    Symbol_Tau                              : 0x00020012,
    Symbol_Upsilon                          : 0x00020013,
    Symbol_Phi                              : 0x00020014,
    Symbol_Chi                              : 0x00020015,
    Symbol_Psi                              : 0x00020016,
    Symbol_Omega                            : 0x00020017,
    //----------------------------------------------------------------------------------------------------------------------
    FractionVertical                        : 0x01000000,
    FractionDiagonal                        : 0x01000001,
    FractionHorizontal                      : 0x01000002,
    FractionSmall                           : 0x01000003,
    FractionDifferential_1                  : 0x01010000,
    FractionDifferential_2                  : 0x01010001,
    FractionDifferential_3                  : 0x01010002,
    FractionDifferential_4                  : 0x01010003,
    FractionPi_2                            : 0x01010004,
    //----------------------------------------------------------------------------------------------------------------------
    ScriptSup                               : 0x02000000,
    ScriptSub                               : 0x02000001,
    ScriptSubSup                            : 0x02000002,
    ScriptSubSupLeft                        : 0x02000003,
    ScriptCustom_1                          : 0x02010000,
    ScriptCustom_2                          : 0x02010001,
    ScriptCustom_3                          : 0x02010002,
    ScriptCustom_4                          : 0x02010003,
    //----------------------------------------------------------------------------------------------------------------------
    RadicalSqrt                             : 0x03000000,
    RadicalRoot_n                           : 0x03000001,
    RadicalRoot_2                           : 0x03000002,
    RadicalRoot_3                           : 0x03000003,
    RadicalCustom_1                         : 0x03010000,
    RadicalCustom_2                         : 0x03010001,
    //----------------------------------------------------------------------------------------------------------------------
    Integral                                : 0x04000000,
    IntegralSubSup                          : 0x04000001,
    IntegralCenterSubSup                    : 0x04000002,
    IntegralDouble                          : 0x04000003,
    IntegralDoubleSubSup                    : 0x04000004,
    IntegralDoubleCenterSubSup              : 0x04000005,
    IntegralTriple                          : 0x04000006,
    IntegralTripleSubSup                    : 0x04000007,
    IntegralTripleCenterSubSup              : 0x04000008,
    IntegralOriented                        : 0x04010000,
    IntegralOrientedSubSup                  : 0x04010001,
    IntegralOrientedCenterSubSup            : 0x04010002,
    IntegralOrientedDouble                  : 0x04010003,
    IntegralOrientedDoubleSubSup            : 0x04010004,
    IntegralOrientedDoubleCenterSubSup      : 0x04010005,
    IntegralOrientedTriple                  : 0x04010006,
    IntegralOrientedTripleSubSup            : 0x04010007,
    IntegralOrientedTripleCenterSubSup      : 0x04010008,
    Integral_dx                             : 0x04020000,
    Integral_dy                             : 0x04020001,
    Integral_dtheta                         : 0x04020002,
    //----------------------------------------------------------------------------------------------------------------------
    LargeOperator_Sum                       : 0x05000000,
    LargeOperator_Sum_CenterSubSup          : 0x05000001,
    LargeOperator_Sum_SubSup                : 0x05000002,
    LargeOperator_Sum_CenterSub             : 0x05000003,
    LargeOperator_Sum_Sub                   : 0x05000004,
    LargeOperator_Prod                      : 0x05010000,
    LargeOperator_Prod_CenterSubSup         : 0x05010001,
    LargeOperator_Prod_SubSup               : 0x05010002,
    LargeOperator_Prod_CenterSub            : 0x05010003,
    LargeOperator_Prod_Sub                  : 0x05010004,
    LargeOperator_CoProd                    : 0x05010005,
    LargeOperator_CoProd_CenterSubSup       : 0x05010006,
    LargeOperator_CoProd_SubSup             : 0x05010007,
    LargeOperator_CoProd_CenterSub          : 0x05010008,
    LargeOperator_CoProd_Sub                : 0x05010009,
    LargeOperator_Union                     : 0x05020000,
    LargeOperator_Union_CenterSubSup        : 0x05020001,
    LargeOperator_Union_SubSup              : 0x05020002,
    LargeOperator_Union_CenterSub           : 0x05020003,
    LargeOperator_Union_Sub                 : 0x05020004,
    LargeOperator_Intersection              : 0x05020005,
    LargeOperator_Intersection_CenterSubSup : 0x05020006,
    LargeOperator_Intersection_SubSup       : 0x05020007,
    LargeOperator_Intersection_CenterSub    : 0x05020008,
    LargeOperator_Intersection_Sub          : 0x05020009,
    LargeOperator_Disjunction               : 0x05030000,
    LargeOperator_Disjunction_CenterSubSup  : 0x05030001,
    LargeOperator_Disjunction_SubSup        : 0x05030002,
    LargeOperator_Disjunction_CenterSub     : 0x05030003,
    LargeOperator_Disjunction_Sub           : 0x05030004,
    LargeOperator_Conjunction               : 0x05030005,
    LargeOperator_Conjunction_CenterSubSup  : 0x05030006,
    LargeOperator_Conjunction_SubSup        : 0x05030007,
    LargeOperator_Conjunction_CenterSub     : 0x05030008,
    LargeOperator_Conjunction_Sub           : 0x05030009,
    LargeOperator_Custom_1                  : 0x05040000,
    LargeOperator_Custom_2                  : 0x05040001,
    LargeOperator_Custom_3                  : 0x05040002,
    LargeOperator_Custom_4                  : 0x05040003,
    LargeOperator_Custom_5                  : 0x05040004,
    //----------------------------------------------------------------------------------------------------------------------
    Bracket_Round                           : 0x06000000,
    Bracket_Square                          : 0x06000001,
    Bracket_Curve                           : 0x06000002,
    Bracket_Angle                           : 0x06000003,
    Bracket_LowLim                          : 0x06000004,
    Bracket_UppLim                          : 0x06000005,
    Bracket_Line                            : 0x06000006,
    Bracket_LineDouble                      : 0x06000007,
    Bracket_Square_OpenOpen                 : 0x06000008,
    Bracket_Square_CloseClose               : 0x06000009,
    Bracket_Square_CloseOpen                : 0x0600000a,
    Bracket_SquareDouble                    : 0x0600000b,
    Bracket_Round_Delimiter_2               : 0x06010000,
    Bracket_Curve_Delimiter_2               : 0x06010001,
    Bracket_Angle_Delimiter_2               : 0x06010002,
    Bracket_Angle_Delimiter_3               : 0x06010003,
    Bracket_Round_OpenNone                  : 0x06020000,
    Bracket_Round_NoneOpen                  : 0x06020001,
    Bracket_Square_OpenNone                 : 0x06020002,
    Bracket_Square_NoneOpen                 : 0x06020003,
    Bracket_Curve_OpenNone                  : 0x06020004,
    Bracket_Curve_NoneOpen                  : 0x06020005,
    Bracket_Angle_OpenNone                  : 0x06020006,
    Bracket_Angle_NoneOpen                  : 0x06020007,
    Bracket_LowLim_OpenNone                 : 0x06020008,
    Bracket_LowLim_NoneNone                 : 0x06020009,
    Bracket_UppLim_OpenNone                 : 0x0602000a,
    Bracket_UppLim_NoneOpen                 : 0x0602000b,
    Bracket_Line_OpenNone                   : 0x0602000c,
    Bracket_Line_NoneOpen                   : 0x0602000d,
    Bracket_LineDouble_OpenNone             : 0x0602000e,
    Bracket_LineDouble_NoneOpen             : 0x0602000f,
    Bracket_SquareDouble_OpenNone           : 0x06020010,
    Bracket_SquareDouble_NoneOpen           : 0x06020011,
    Bracket_Custom_1                        : 0x06030000,
    Bracket_Custom_2                        : 0x06030001,
    Bracket_Custom_3                        : 0x06030002,
    Bracket_Custom_4                        : 0x06030003,
    Bracket_Custom_5                        : 0x06040000,
    Bracket_Custom_6                        : 0x06040001,
    Bracket_Custom_7                        : 0x06040002,
    //----------------------------------------------------------------------------------------------------------------------
    Function_Sin                            : 0x07000000,
    Function_Cos                            : 0x07000001,
    Function_Tan                            : 0x07000002,
    Function_Csc                            : 0x07000003,
    Function_Sec                            : 0x07000004,
    Function_Cot                            : 0x07000005,
    Function_1_Sin                          : 0x07010000,
    Function_1_Cos                          : 0x07010001,
    Function_1_Tan                          : 0x07010002,
    Function_1_Csc                          : 0x07010003,
    Function_1_Sec                          : 0x07010004,
    Function_1_Cot                          : 0x07010005,
    Function_Sinh                           : 0x07020000,
    Function_Cosh                           : 0x07020001,
    Function_Tanh                           : 0x07020002,
    Function_Csch                           : 0x07020003,
    Function_Sech                           : 0x07020004,
    Function_Coth                           : 0x07020005,
    Function_1_Sinh                         : 0x07030000,
    Function_1_Cosh                         : 0x07030001,
    Function_1_Tanh                         : 0x07030002,
    Function_1_Csch                         : 0x07030003,
    Function_1_Sech                         : 0x07030004,
    Function_1_Coth                         : 0x07030005,
    Function_Custom_1                       : 0x07040000,
    Function_Custom_2                       : 0x07040001,
    Function_Custom_3                       : 0x07040002,
    //----------------------------------------------------------------------------------------------------------------------
    Accent_Dot                              : 0x08000000,
    Accent_DDot                             : 0x08000001,
    Accent_DDDot                            : 0x08000002,
    Accent_Hat                              : 0x08000003,
    Accent_Check                            : 0x08000004,
    Accent_Accent                           : 0x08000005,
    Accent_Grave                            : 0x08000006,
    Accent_Smile                            : 0x08000007,
    Accent_Tilde                            : 0x08000008,
    Accent_Bar                              : 0x08000009,
    Accent_DoubleBar                        : 0x0800000a,
    Accent_CurveBracketTop                  : 0x0800000b,
    Accent_CurveBracketBot                  : 0x0800000c,
    Accent_GroupTop                         : 0x0800000d,
    Accent_GroupBot                         : 0x0800000e,
    Accent_ArrowL                           : 0x0800000f,
    Accent_ArrowR                           : 0x08000010,
    Accent_ArrowD                           : 0x08000011,
    Accent_HarpoonL                         : 0x08000012,
    Accent_HarpoonR                         : 0x08000013,
    Accent_BorderBox                        : 0x08010000,
    Accent_BorderBoxCustom                  : 0x08010001,
    Accent_BarTop                           : 0x08020000,
    Accent_BarBot                           : 0x08020001,
    Accent_Custom_1                         : 0x08030000,
    Accent_Custom_2                         : 0x08030001,
    Accent_Custom_3                         : 0x08030002,
    //----------------------------------------------------------------------------------------------------------------------
    LimitLog_LogBase                        : 0x09000000,
    LimitLog_Log                            : 0x09000001,
    LimitLog_Lim                            : 0x09000002,
    LimitLog_Min                            : 0x09000003,
    LimitLog_Max                            : 0x09000004,
    LimitLog_Ln                             : 0x09000005,
    LimitLog_Custom_1                       : 0x09010000,
    LimitLog_Custom_2                       : 0x09010001,
    //----------------------------------------------------------------------------------------------------------------------
    Operator_ColonEquals                    : 0x0a000000,
    Operator_EqualsEquals                   : 0x0a000001,
    Operator_PlusEquals                     : 0x0a000002,
    Operator_MinusEquals                    : 0x0a000003,
    Operator_Definition                     : 0x0a000004,
    Operator_UnitOfMeasure                  : 0x0a000005,
    Operator_DeltaEquals                    : 0x0a000006,
    Operator_ArrowL_Top                     : 0x0a010000,
    Operator_ArrowR_Top                     : 0x0a010001,
    Operator_ArrowL_Bot                     : 0x0a010002,
    Operator_ArrowR_Bot                     : 0x0a010003,
    Operator_DoubleArrowL_Top               : 0x0a010004,
    Operator_DoubleArrowR_Top               : 0x0a010005,
    Operator_DoubleArrowL_Bot               : 0x0a010006,
    Operator_DoubleArrowR_Bot               : 0x0a010007,
    Operator_ArrowD_Top                     : 0x0a010008,
    Operator_ArrowD_Bot                     : 0x0a010009,
    Operator_DoubleArrowD_Top               : 0x0a01000a,
    Operator_DoubleArrowD_Bot               : 0x0a01000b,
    Operator_Custom_1                       : 0x0a020000,
    Operator_Custom_2                       : 0x0a020001,
    //----------------------------------------------------------------------------------------------------------------------
    Matrix_1_2                              : 0x0b000000,
    Matrix_2_1                              : 0x0b000001,
    Matrix_1_3                              : 0x0b000002,
    Matrix_3_1                              : 0x0b000003,
    Matrix_2_2                              : 0x0b000004,
    Matrix_2_3                              : 0x0b000005,
    Matrix_3_2                              : 0x0b000006,
    Matrix_3_3                              : 0x0b000007,
    Matrix_Dots_Center                      : 0x0b010000,
    Matrix_Dots_Baseline                    : 0x0b010001,
    Matrix_Dots_Vertical                    : 0x0b010002,
    Matrix_Dots_Diagonal                    : 0x0b010003,
    Matrix_Identity_2                       : 0x0b020000,
    Matrix_Identity_2_NoZeros               : 0x0b020001,
    Matrix_Identity_3                       : 0x0b020002,
    Matrix_Identity_3_NoZeros               : 0x0b020003,
    Matrix_2_2_RoundBracket                 : 0x0b030000,
    Matrix_2_2_SquareBracket                : 0x0b030001,
    Matrix_2_2_LineBracket                  : 0x0b030002,
    Matrix_2_2_DLineBracket                 : 0x0b030003,
    Matrix_Flat_Round                       : 0x0b040000,
    Matrix_Flat_Square                      : 0x0b040001,
    //----------------------------------------------------------------------------------------------------------------------
    Default_Text                            : 0x0c000000
};

function CRPI()
{
    this.bDecreasedComp            = false;
    this.bInline                   = false;
    this.bChangeInline             = false;
    this.bNaryInline               = false; /*для CDegreeSupSub внутри N-арного оператора, этот флаг необходим, чтобы итераторы максимально близко друг к другу расположить*/
    this.bEqArray                  = false; /*для амперсанда*/
    this.bMathFunc                 = false;
    this.bRecalcCtrPrp             = false; // пересчет ctrPrp нужен, когда на Undo и тп изменился размер первого Run, а ctrPrp уже для мат объектов пересчитались
    this.bCorrect_ConvertFontSize  = false;
    this.bSmallFraction            = false;
}
CRPI.prototype.MergeMathInfo = function(MathInfo)
{
    this.bInline                   = MathInfo.bInline || (MathInfo.bInternalRanges == true && MathInfo.bStartRanges == false);
    this.bRecalcCtrPrp             = MathInfo.bRecalcCtrPrp;
    this.bChangeInline             = MathInfo.bChangeInline;
    this.bCorrect_ConvertFontSize  = MathInfo.bCorrect_ConvertFontSize;
};

function CMathPointInfo()
{
    this.x    = 0;
    this.y    = 0;

    this.bEven      = true;
    this.CurrPoint  = 0;

    this.InfoPoints = {};
}
CMathPointInfo.prototype.SetInfoPoints = function(InfoPoints)
{
    this.InfoPoints.GWidths       = InfoPoints.GWidths;
    this.InfoPoints.GPoints       = InfoPoints.GPoints;
    this.InfoPoints.ContentPoints = InfoPoints.ContentPoints.Widths;
    this.InfoPoints.GMaxDimWidths = InfoPoints.GMaxDimWidths;
};
CMathPointInfo.prototype.NextAlignRange = function()
{
    if(this.bEven)
        this.bEven = false;
    else
    {
        this.CurrPoint++;
        this.bEven = true;
    }
};
CMathPointInfo.prototype.GetAlign = function()
{
    var align = 0;

    if(this.bEven)
    {
        var alignEven, alignGeneral, alignOdd;

        var Len   = this.InfoPoints.ContentPoints.length,
            Point = this.InfoPoints.ContentPoints[this.CurrPoint];

        var GWidth = this.InfoPoints.GWidths[this.CurrPoint],
            GPoint = this.InfoPoints.GPoints[this.CurrPoint];

        if(this.CurrPoint == Len - 1 && Point.odd == -1) // то есть последняя точка четная, выравнивание по центру
        {
            var GMaxDimWidth = this.InfoPoints.GMaxDimWidths[this.CurrPoint];

            alignGeneral = (GMaxDimWidth - Point.even)/2;
            alignEven = 0;
        }
        else
        {
            alignGeneral = (GWidth - GPoint.even - GPoint.odd)/2;
            alignEven = GPoint.even - Point.even;
        }

        if(this.CurrPoint > 0)
        {
            var PrevGenPoint = this.InfoPoints.GPoints[this.CurrPoint-1],
                PrevGenWidth = this.InfoPoints.GWidths[this.CurrPoint-1],
                PrevPoint    = this.InfoPoints.ContentPoints[this.CurrPoint-1];

            var alignPrevGen = (PrevGenWidth - PrevGenPoint.even - PrevGenPoint.odd)/2;
            alignOdd = alignPrevGen +  PrevGenPoint.odd - PrevPoint.odd;
        }
        else
            alignOdd = 0;

        align = alignGeneral + alignEven + alignOdd;
    }

    return align;
};

function CInfoPoints()
{
    this.GWidths       = null;
    this.GPoints       = null;
    this.GMaxDimWidths = null;
    this.ContentPoints = new AmperWidths();
}
CInfoPoints.prototype.SetDefault = function()
{
    this.GWidths       = null;
    this.GPoints       = null;
    this.GMaxDimWidths = null;
    this.ContentPoints.SetDefault();
};

function CMathPosInfo()
{
    this.CurRange = -1;
    this.CurLine  = -1;

    this.DispositionOpers = null;
}

function CMathPosition()
{
    this.x  = 0;
    this.y  = 0;
}
CMathPosition.prototype.Set = function(Pos)
{
    this.x = Pos.x;
    this.y = Pos.y;
};

function AmperWidths()
{
    this.bEven     = true; // является ли текущая точка нечетной
    this.Widths    = [];
}
AmperWidths.prototype.UpdatePoint = function(value)
{
    var len = this.Widths.length;

    if(len == 0)
    {
        // дефолтное значение bEven true, для случая если первый элемент в контенте будет Ampersand
        var NewPoint = new CMathPoint();
        NewPoint.even = value;
        this.Widths.push(NewPoint);
        this.bEven = true;
    }
    else
    {
        if(this.bEven)
            this.Widths[len-1].even += value;
        else
            this.Widths[len-1].odd += value;
    }

};
AmperWidths.prototype.AddNewAlignRange = function()
{
    var len = this.Widths.length;

    if(!this.bEven || len == 0)
    {
        var NewPoint = new CMathPoint();
        NewPoint.even = 0;
        this.Widths.push(NewPoint);
    }

    if(this.bEven)
    {
        len = this.Widths.length;
        this.Widths[len-1].odd = 0;
    }


    this.bEven = !this.bEven;

};
AmperWidths.prototype.SetDefault = function()
{
    this.bEven         = true;
    this.Widths.length = 0;
};


function CGeneralObjectGaps(Left, Right)
{
    this.left  = Left;
    this.right = Right;
}

function CGaps(oSign, oEqual, oZeroOper, oLett)
{
    this.sign = oSign;
    this.equal = oEqual;
    this.zeroOper = oZeroOper;
    this.letters = oLett;
}

function CCoeffGaps()
{
    var LeftSign  = new CGaps(0.52, 0.26, 0, 0.52),
        RightSign = new CGaps(0.49, 0, 0, 0.49);

    this.Sign = new CGeneralObjectGaps(LeftSign, RightSign);

    var LeftMult  = new CGaps(0, 0, 0, 0.46),
        RightMult = new CGaps(0, 0, 0, 0.49);

    this.Mult = new CGeneralObjectGaps(LeftMult, RightMult);

    var LeftEqual  = new CGaps(0, 0, 0, 0.7),
        RightEqual = new CGaps(0, 0, 0, 0.5);

    this.Equal = new CGeneralObjectGaps(LeftEqual, RightEqual);

    var LeftDefault  = new CGaps(0, 0, 0, 0),
        RightDefault = new CGaps(0, 0, 0, 0);

    this.Default = new CGeneralObjectGaps(LeftDefault, RightDefault);
}
CCoeffGaps.prototype =
{
    getCoeff: function(codeCurr, codeLR , direct) // obj - либо codeChar, либо мат объект
    {
        var operator = null;

        if(this.checkEqualSign(codeCurr))
            operator = this.Equal;
        else if(this.checkOperSign(codeCurr))
            operator = this.Sign;
        else if(codeCurr == 0x2A)
            operator = this.Mult;
        else
            operator = this.Default;

        var part = direct == -1 ? operator.left : operator.right;

        var coeff = 0;
        if(codeLR == -1) // мат объект
            coeff = part.letters;
        else if(this.checkOperSign(codeLR))
            coeff = part.sign;
        else if(this.checkEqualSign(codeLR))
            coeff = part.equal;
        else if(this.checkZeroSign(codeLR, direct))
            coeff = part.zeroOper;
        else
            coeff = part.letters;

        return coeff;
    },
    checkOperSign: function(code) // "+", "-", "±", "∓", "×", "÷"
    {
        var PLUS            = 0x2B,
            MINUS           = 0x2D,
            PLUS_MINUS      = 0xB1,
            MINUS_PLUS      = 0x2213,
            MULTIPLICATION  = 0xD7,
            DIVISION        = 0xF7;

        return code == PLUS || code == MINUS || code == PLUS_MINUS || code == MINUS_PLUS || code == MULTIPLICATION || code == DIVISION;
    },
    checkEqualSign: function(code)
    {
        var COMPARE       = code == 0x3C || code == 0x3E; // LESS, GREATER
        var ARROWS        = (code >= 0x2190 && code <= 0x21B3) || (code == 0x21B6) || (code == 0x21B7) || (code >= 0x21BA && code <= 0x21E9) || (code >=0x21F4 && code <= 0x21FF);
        var INTERSECTION  = code >= 0x2223 && code <= 0x222A;
        var EQUALS        = code == 0x3D || (code >= 0x2234 && code <= 0x22BD) || (code >= 0x22C4 && code <= 0x22FF);
        var ARR_FISHES    = (code >= 0x27DA && code <= 0x27E5) || (code >= 0x27EC && code <= 0x297F);
        var TRIANGLE_SYMB = code >= 0x29CE && code <= 0x29D7;
        var OTH_SYMB      = code == 0x29DF || (code >= 0x29E1 && code <= 0x29E7) || (code >= 0x29F4 && code <= 0x29F8) || (code >= 0x2A22 && code <= 0x2AF0) || (code >= 0x2AF2 && code <= 0x2AFB) || code == 0x2AFD || code == 0x2AFE;


        return COMPARE || ARROWS || INTERSECTION || EQUALS || ARR_FISHES || TRIANGLE_SYMB || OTH_SYMB;
    },
    checkZeroSign: function(code, direct) // "*", "/", "\"
    {
        var MULT     = 0x2A,
            DIVISION = 0x2F,
            B_SLASH  = 0x5C;

        var bOper = code == MULT || code == DIVISION || code == B_SLASH;
        var bLeftBracket = direct == -1 && (code == 0x28 || code == 0x5B || code == 0x7B);
        var bRightBracket = direct == 1 && (code == 0x29 || code == 0x5D || code == 0x7D);


        return bOper || bLeftBracket || bRightBracket;
    }
};

var COEFF_GAPS = new CCoeffGaps();

function CMathArgSize()
{
    this.value       = undefined;
}
CMathArgSize.prototype =
{
    Decrease: function()
    {
        if(this.value == undefined)
            this.value = 0;

        if( this.value > -2 )
            this.value--;

        return this.value;
    },
    Increase: function()
    {
        if(this.value == undefined)
            this.value = 0;

        if(this.value < 2)
            this.value++;

        return this.value;
    },
    Set: function(ArgSize)
    {
        this.value = ArgSize.value;
    },
    GetValue: function()
    {
        return this.value;
    },
    SetValue: function(val)
    {
        if(val === null || val === undefined)
            this.value = undefined;
        else if(val < - 2)
            this.value = -2;
        else if(val > 2)
            this.value = 2;
        else
            this.value = val;
    },
    Copy: function()
    {
        var ArgSize = new CMathArgSize();
        ArgSize.value = this.value;

        return ArgSize;
    },
    Merge: function(ArgSize)
    {
        if(this.value == undefined)
            this.value = 0;

        if(ArgSize.value == undefined)
            ArgSize.value = 0;

        this.SetValue(this.value + ArgSize.value);
    },
    Can_Decrease: function()
    {
        return this.value !== -2;
    },
    Can_Increase: function()
    {
        return this.value == -1 || this.value == -2;
    },
    Can_SimpleIncrease: function()
    {
        return this.value !== 2;
    },
    Is_Init: function()
    {
        return this.value !== undefined;
    },
    Write_ToBinary: function(Writer)
    {
        if(this.value == undefined)
        {
            Writer.WriteBool(true);
        }
        else
        {
            Writer.WriteBool(false);
            Writer.WriteLong(this.value);
        }
    },
    Read_FromBinary: function(Reader)
    {
        if(Reader.GetBool() == false)
        {
            this.value = Reader.GetLong();
        }
        else
        {
            this.value = undefined;
        }
    }
};

function CMathGapsInfo(argSize)
{
    this.argSize = argSize; // argSize выставляем один раз для всего контента

    this.Left    = null;    // элемент слева
    this.Current = null;    // текущий элемент

    this.LeftFontSize    = null;
    this.CurrentFontSize = null;
    this.bUpdate         = false;
}
CMathGapsInfo.prototype =
{
    setGaps: function(Current, CurrentFontSize)
    {
        this.updateCurrentObject(Current, CurrentFontSize);
        this.updateGaps();
    },
    updateCurrentObject: function(Current, CurrentFontSize)
    {
        this.Left = this.Current;
        this.LeftFontSize = this.CurrentFontSize;

        this.Current = Current;
        this.CurrentFontSize = CurrentFontSize;
    },
    updateGaps: function()
    {
        if(this.argSize < 0)
        {
            this.Current.GapLeft = 0;

            if(this.Left !== null)
                this.Left.GapRight = 0;
        }
        else
        {
            var leftCoeff = 0,  /// for Current Object
                rightCoeff = 0; /// for Left Object

            var leftCode;


            if(this.Current.IsMathText())
            {
                var currCode = this.Current.getCodeChr();

                if(this.Left !== null)
                {
                    if(this.Left.Type == para_Math_Composition)
                    {
                        rightCoeff = this.getGapsMComp(this.Left, 1);
                        leftCoeff = COEFF_GAPS.getCoeff(currCode, -1, -1);

                        if(leftCoeff > rightCoeff)
                            leftCoeff -= rightCoeff;
                    }
                    else if(this.Left.IsMathText())
                    {
                        leftCode = this.Left.getCodeChr();
                        leftCoeff = COEFF_GAPS.getCoeff(currCode, leftCode, -1);
                        rightCoeff = COEFF_GAPS.getCoeff(leftCode, currCode, 1);
                    }

                }
                else
                    this.Current.GapLeft = 0;
            }
            else if(this.Current.Type == para_Math_Composition)
            {
                leftCoeff = this.getGapsMComp(this.Current, -1);

                if(this.Left !== null)
                {
                    if(this.Left.Type == para_Math_Composition)
                    {
                        rightCoeff = this.getGapsMComp(this.Left, 1);

                        if(rightCoeff/2 > leftCoeff)
                            rightCoeff -= leftCoeff;
                        else
                            rightCoeff /= 2;

                        if(leftCoeff < rightCoeff/2)
                            leftCoeff = rightCoeff/2;
                        else
                            leftCoeff -= rightCoeff/2;
                    }
                    else if(this.Left.IsMathText())
                    {
                        leftCode = this.Left.getCodeChr();
                        rightCoeff = COEFF_GAPS.getCoeff(leftCode, -1, 1);
                        if(rightCoeff > leftCoeff)
                            rightCoeff -= leftCoeff;
                    }
                }
                else
                    leftCoeff = 0;
            }

            var LGapSign = 0.1513*this.CurrentFontSize;
            this.Current.GapLeft = (leftCoeff*LGapSign*100 | 0)/100; // если ни один случай не выполнился, выставляем "нулевые" gaps (default): необходимо, если что-то удалили и объект стал первый или последним в контенте

            if(this.Left !== null)
            {
                var RGapSign = 0.1513*this.LeftFontSize;
                this.Left.GapRight = (rightCoeff*RGapSign*100 | 0)/100;
            }
        }
    },
    getGapsMComp: function(MComp, direct)
    {
        var kind = MComp.kind;
        var checkGap = this.checkGapKind(MComp);

        var bNeedGap = !checkGap.bEmptyGaps && !checkGap.bChildGaps;

        var coeffLeft  = 0.001,
            coeffRight = 0; // for checkGap.bEmptyGaps

        //var bDegree = kind == MATH_DEGREE || kind == MATH_DEGREESubSup;
        var bDegree = kind == MATH_DEGREE;

        if(checkGap.bChildGaps)
        {
            if(bDegree)
            {
                coeffLeft  = 0.03;

                if(MComp.IsPlhIterator())
                    coeffRight = 0.12;
                else
                    coeffRight = 0.16;
            }

            var gapsChild = MComp.getGapsInside(this);

            coeffLeft  = coeffLeft  < gapsChild.left  ? gapsChild.left  : coeffLeft;
            coeffRight = coeffRight < gapsChild.right ? gapsChild.right : coeffRight;
        }
        else if(bNeedGap)
        {
            coeffLeft = 0.4;
            coeffRight = 0.3;
        }


        return direct == -1 ? coeffLeft : coeffRight;
    },
    checkGapKind: function(Comp)
    {
        var kind       = Comp.kind;

        var bEmptyGaps = kind == MATH_MATRIX /*|| (kind == MATH_DELIMITER && Comp.Is_EmptyGaps())*/,
            bChildGaps = kind == MATH_DEGREE || kind == MATH_DEGREESubSup || kind == MATH_ACCENT || kind == MATH_RADICAL || kind == MATH_LIMIT || kind == MATH_BORDER_BOX;


        return  {bEmptyGaps: bEmptyGaps, bChildGaps: bChildGaps};
    }
};

function CMPrp()
{
    this.sty      = undefined;
    this.scr      = undefined;
    this.nor      = undefined;

    this.aln      = undefined;
    this.brk      = undefined;
    this.lit      = undefined;

    // Default
    /*this.sty    = STY_ITALIC;
    this.scr      = TXT_ROMAN;

    this.nor      = false;

    this.aln      = false;
    this.brk      = false;
    this.lit      = false;*/

    // TXT_NORMAL
    // если normal == false, то берем TextPrp отсюда (в wRunPrp bold/italic не учитываем, выставляем отсюда)
    // если normal == true, то их Word не учитывает и берет TextPr из wRunPrp

    // TXT_PLAIN
    // если plain == true
    // буквы берутся обычные, не специальные для Cambria Math : то есть как для TXT_NORMAL
    // отличие от TXT_NORMAL w:rPrp в этом случае не учитываются !

}
CMPrp.prototype =
{
    Is_Empty: function() {
        return undefined === this.sty && undefined === this.scr && undefined === this.nor && undefined === this.aln &&
            undefined === this.brk && undefined === this.lit;
    },
    Set_Pr: function(Pr)
    {
        if(Pr.sty !== undefined)
            this.sty = Pr.sty;

        if(Pr.scr !== undefined)
            this.scr = Pr.scr;

        if(Pr.nor !== undefined)
            this.nor = Pr.nor;

        if(Pr.aln !== undefined)
            this.aln = Pr.aln;

        if(Pr.lit !== undefined)
            this.lit = Pr.lit;

        if(Pr.brk !== undefined)
        {
            this.brk = new CMathBreak();
            this.brk.Set_FromObject(Pr.brk);
        }
    },
    GetTxtPrp: function()
    {
        var textPrp = new CTextPr();

        if(this.sty == undefined)
        {
            textPrp.Italic = true;
            textPrp.Bold   = false;
        }
        else
        {
            textPrp.Italic = this.sty == STY_BI || this.sty == STY_ITALIC;
            textPrp.Bold   = this.sty == STY_BI || this.sty == STY_BOLD;
        }


        return textPrp;
    },
    Copy: function()
    {
        var NewMPrp = new CMPrp();
        
        NewMPrp.aln      = this.aln;
        NewMPrp.lit      = this.lit;
        NewMPrp.nor      = this.nor;
        NewMPrp.sty      = this.sty;
        NewMPrp.scr      = this.scr;

        if(this.brk !== undefined)
            NewMPrp.brk = this.brk.Copy();
        
        return NewMPrp;
    },
    IsBreak: function()
    {
        return this.brk !== undefined;
    },
    Get_AlignBrk: function()
    {
        return this.brk !== undefined ? this.brk.Get_AlignBrk() : null;
    },
    Get_AlnAt: function()
    {
        return this.brk != undefined ? this.brk.Get_AlnAt() : undefined;
    },
    GetCompiled_ScrStyles : function()
    {
        var nor = this.nor == undefined ? false : this.nor;
        var scr = this.scr == undefined ? TXT_ROMAN : this.scr;
        var sty = this.sty == undefined ? STY_ITALIC : this.sty;

        return {nor: nor, scr: scr, sty: sty};
    },
    SetStyle: function(Bold, Italic) /// из ctrPrp получить style для MathPrp
    {
        if(Bold == true && Italic == true)
            this.sty = STY_BI;
        else if(Italic == true)
            this.sty = STY_ITALIC;
        else if(Bold == true)
            this.sty = STY_BOLD;
        else if(Bold == false && Italic == false)
            this.sty = STY_PLAIN;
        else
            this.sty = undefined;
    },
    GetBoldItalic: function()
    {
        var Object =
        {
            Italic: undefined,
            Bold:   undefined
        };

        if(this.sty == STY_BI)
        {
            Object.Bold   = true;

        }
        else if(this.sty == STY_BOLD)
        {
            Object.Bold   = true;
            Object.Italic = false;
        }

        return Object;
    },
    Displace_Break: function(isForward)
    {
        if(this.brk !== undefined)
        {
            this.brk.Displace(isForward);
        }
    },
    Apply_AlnAt: function(alnAt)
    {
        if(this.brk !== undefined)
        {
            this.brk.Apply_AlnAt(alnAt);
        }
    },
    Insert_ForcedBreak: function(AlnAt)
    {
        if(this.brk == undefined)
            this.brk = new CMathBreak();

        this.brk.Apply_AlnAt(AlnAt);
    },
    Delete_ForcedBreak: function()
    {
        this.brk = undefined;
    }
};
CMPrp.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);

	var Flags = 0;
	if (undefined != this.aln)
	{
		Writer.WriteBool(this.aln);
		Flags |= 1;
	}
	if (undefined != this.brk)
	{
		this.brk.Write_ToBinary(Writer);
		Flags |= 2;
	}
	if (undefined != this.lit)
	{
		Writer.WriteBool(this.lit);
		Flags |= 4;
	}
	if (undefined != this.nor)
	{
		Writer.WriteBool(this.nor);
		Flags |= 8;
	}
	if (undefined != this.scr)
	{
		Writer.WriteLong(this.scr);
		Flags |= 16;
	}
	if (undefined != this.sty)
	{
		Writer.WriteLong(this.sty);
		Flags |= 32;
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CMPrp.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	if (Flags & 1)
		this.aln = Reader.GetBool();

	if (Flags & 2)
	{
		this.brk = new CMathBreak();
		this.brk.Read_FromBinary(Reader);
	}

	if (Flags & 4)
		this.lit = Reader.GetBool();

	if (Flags & 8)
		this.nor = Reader.GetBool();

	if (Flags & 16)
		this.scr = Reader.GetLong();

	if (Flags & 32)
		this.sty = Reader.GetLong();
};

/**
 *
 * @constructor
 * @extends {CParagraphContentWithParagraphLikeContent}
 */
function CMathContent()
{
	CParagraphContentWithParagraphLikeContent.call(this);

	this.Id = AscCommon.g_oIdCounter.Get_NewId();		

    this.Content = []; // array of mathElem

    this.Type = para_Math_Content;
    this.CurPos = 0;
    this.pos    = new CMathPosition();   // относительная позиция

    //  Properties
    this.ParaMath       = null;
    this.ArgSize        = new CMathArgSize();
    this.Compiled_ArgSz = new CMathArgSize();

    // for EqArray
    this.InfoPoints = new CInfoPoints();
    ///////////////

    this.Bounds = new CMathBounds();

    this.plhHide        = false;
    this.bRoot          = false;
    this.bOneLine  = false;

    //////////////////

    this.Selection =
    {
        StartPos:  0,
        EndPos:    0,
        Use:    false
    };

    this.RecalcInfo =
    {
        TextPr:             true,
        bEqArray:          false,
        bChangeInfoPoints:  false,
        Measure:            true
    };

    this.NearPosArray = [];
    this.ParentElement = null;

    this.size = new CMathSize();

	this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)
	
	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
    AscCommon.g_oTableId.Add( this, this.Id );
}
CMathContent.prototype = Object.create(CParagraphContentWithParagraphLikeContent.prototype);
CMathContent.prototype.constructor = CMathContent;
CMathContent.prototype.init = function()
{

};
CMathContent.prototype.addElementToContent = function(obj)
{
    this.Internal_Content_Add(this.Content.length, obj, false);
    this.CurPos = this.Content.length-1;
};
CMathContent.prototype.SetPlaceholder = function()
{
	this.Remove_FromContent(0, this.Content.length);

    var oRun = new ParaRun(null, true);
    oRun.AddMathPlaceholder();

    this.addElementToContent(oRun);

    this.RemoveSelection();
    this.MoveCursorToEndPos();
};
//////////////////////////////////////
CMathContent.prototype.PreRecalc = function(Parent, ParaMath, ArgSize, RPI)
{
    this.ParaMath = ParaMath;
    if(Parent !== null)
    {
        this.bRoot = false;
        this.Parent = Parent;
    }

    if(ArgSize !== null && ArgSize !== undefined)
    {
        this.Compiled_ArgSz.value = this.ArgSize.value;
        this.Compiled_ArgSz.Merge(ArgSize);
    }

    var lng = this.Content.length;

    var GapsInfo = new CMathGapsInfo(this.Compiled_ArgSz.value);

    if(!this.bRoot)
        this.RecalcInfo.bEqArray = this.Parent.IsEqArray();

    for(var pos = 0; pos < lng; pos++)
    {
        if(this.Content[pos].Type == para_Math_Composition)
        {
            this.Content[pos].PreRecalc(this, ParaMath, this.Compiled_ArgSz, RPI, GapsInfo);
        }
        else if(this.Content[pos].Type == para_Math_Run)
            this.Content[pos].Math_PreRecalc(this, ParaMath, this.Compiled_ArgSz, RPI, GapsInfo);
    }

    if(GapsInfo.Current !== null)
        GapsInfo.Current.GapRight = 0;
};
CMathContent.prototype.Math_UpdateGaps = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var GapsInfo = new CMathGapsInfo(this.Compiled_ArgSz.value);

    if(StartPos !== undefined && EndPos !== undefined && CurLine < this.protected_GetLinesCount())
    {
        if(CurLine > 0 && StartPos !== EndPos) // выставим объект, который будет Left для первого элемента в текущей строке
        {
            var EndPosPrev   = this.protected_GetRangeEndPos(CurLine - 1, CurRange);
            this.Content[EndPosPrev].UpdLastElementForGaps(_CurLine - 1, _CurRange, GapsInfo);
        }


        for(var Pos = StartPos; Pos <= EndPos; Pos++)
        {
            this.Content[Pos].Math_UpdateGaps(_CurLine, _CurRange, GapsInfo);
        }
    }
};
CMathContent.prototype.IsEqArray = function()
{
    return this.RecalcInfo.bEqArray;
};
CMathContent.prototype.Get_WidthPoints = function()
{
    return this.InfoPoints.ContentPoints;
};
CMathContent.prototype.ShiftPage = function(Dx)
{
    this.Bounds.ShiftPage(Dx);

    for(var Pos = 0; Pos < this.Content.length; Pos++)
    {
        if(this.Content[Pos].Type === para_Math_Composition)
            this.Content[Pos].ShiftPage(Dx);
    }
};
CMathContent.prototype.Get_CompiledArgSize = function()
{
    return this.Compiled_ArgSz;
};
CMathContent.prototype.getGapsInside = function(GapsInfo) // учитываем gaps внутренних объектов
{
    var gaps = {left: 0, right: 0};
    var bFirstComp = false,
        bLastComp = false;

    var len = this.Content.length;

    if(len > 1)
    {
        var bFRunEmpty = this.Content[0].Is_Empty();
        bFirstComp = bFRunEmpty && this.Content[1].Type == para_Math_Composition; // первый всегда идет Run

        var bLastRunEmpty = this.Content[len - 1].Is_Empty(); // т.к. после мат. объекта стоит пустой Run
        bLastComp = bLastRunEmpty && this.Content[len - 2].Type == para_Math_Composition;
    }

    var checkGap;

    if(bFirstComp)
    {
        checkGap = GapsInfo.checkGapKind(this.Content[1]);

        if(!checkGap.bChildGaps)
        {
            gaps.left = GapsInfo.getGapsMComp(this.Content[1], -1);
        }
    }

    if(bLastComp)
    {
        checkGap = GapsInfo.checkGapKind(this.Content[len - 1]);

        if(!checkGap.bChildGaps)
        {
            gaps.right = GapsInfo.getGapsMComp(this.Content[len - 2], 1);
        }
    }

    return gaps;
};
CMathContent.prototype.draw = function(x, y, pGraphics, PDSE)
{
    var StartPos, EndPos;

    if(this.bRoot)
    {
        var CurLine  = PDSE.Line - this.StartLine;
        var CurRange = ( 0 === CurLine ? PDSE.Range - this.StartRange : PDSE.Range );

        StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
        EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);
    }
    else
    {
        StartPos = 0;
        EndPos   = this.Content.length - 1;
    }

    var bHidePlh = this.plhHide && this.IsPlaceholder();

    if( !bHidePlh )
    {
        for(var i = StartPos; i <= EndPos;i++)
        {
            if(this.Content[i].Type == para_Math_Composition)
            {
                this.Content[i].draw(x, y, pGraphics, PDSE);
            }
            else
                this.Content[i].Draw_Elements(PDSE);
        }
    }
};
CMathContent.prototype.Draw_Elements = function(PDSE)
{
    var StartPos, EndPos;

    if(this.protected_GetLinesCount() > 0)
    {
        var CurLine  = PDSE.Line - this.StartLine;
        var CurRange = ( 0 === CurLine ? PDSE.Range - this.StartRange : PDSE.Range );

        StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
        EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);
    }
    else
    {
        StartPos = 0;
        EndPos   = this.Content.length - 1;
    }

    var bHidePlh = this.plhHide && this.IsPlaceholder();

    if( !bHidePlh )
    {
        for(var CurPos = StartPos; CurPos <= EndPos;CurPos++)
            this.Content[CurPos].Draw_Elements(PDSE);
    }

};
CMathContent.prototype.setCtrPrp = function()
{

};
CMathContent.prototype.Is_InclineLetter = function()
{
    var result = false;

    if(this.Content.length == 1)
        result = this.Content[0].Math_Is_InclineLetter();

    return result;
};
CMathContent.prototype.IsPlaceholder = function()
{
    var bPlh = false;
    if(this.Content.length == 1)
        bPlh  = this.Content[0].IsPlaceholder();

    return bPlh;
};
CMathContent.prototype.Can_GetSelection = function()
{
    var bPlh = false;

    if(this.Content.length == 1)
        bPlh  = this.Content[0].IsPlaceholder();

    return bPlh || this.bRoot == false;
};
CMathContent.prototype.IsJustDraw = function()
{
    return false;
};
CMathContent.prototype.ApplyPoints = function(WidthsPoints, Points, MaxDimWidths)
{
    this.InfoPoints.GWidths       = WidthsPoints;
    this.InfoPoints.GPoints       = Points;
    this.InfoPoints.GMaxDimWidths = MaxDimWidths;
    // точки выравнивания данного контента содержатся в ContentPoints

    var PosInfo = new CMathPointInfo();
    PosInfo.SetInfoPoints(this.InfoPoints);

    this.size.width = 0;

    for(var i = 0 ; i < this.Content.length; i++)
    {
        if(this.Content[i].Type === para_Math_Run)
        {
            this.Content[i].ApplyPoints(PosInfo);
        }

        this.size.width += this.Content[i].size.width;
    }

    this.Bounds.SetWidth(0, 0, this.size.width);
};
CMathContent.prototype.UpdateBoundsPosInfo = function(PRSA, _CurLine, _CurRange, _CurPage)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    this.Bounds.SetGenPos(CurLine, CurRange, PRSA);
    this.Bounds.SetPage(CurLine, CurRange, _CurPage);

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        if(this.Content[Pos].Type == para_Math_Composition)
            this.Content[Pos].UpdateBoundsPosInfo(PRSA, _CurLine, _CurRange, _CurPage);
    }

};
CMathContent.prototype.setPosition = function(pos, PosInfo)
{
    var Line  = PosInfo.CurLine,
        Range = PosInfo.CurRange;

    var CurLine  = Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? Range - this.StartRange : Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    if(this.RecalcInfo.bEqArray)
    {
        var PosInfoEqq = new CMathPointInfo();
        PosInfoEqq.SetInfoPoints(this.InfoPoints);

        pos.x += PosInfoEqq.GetAlign();
    }

    this.pos.x = pos.x;
    this.pos.y = pos.y;

    this.Bounds.SetPos(CurLine, CurRange, this.pos);

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        if(this.Content[Pos].Type == para_Math_Run)
            this.Content[Pos].Math_SetPosition(pos, PosInfo);
        else
            this.Content[Pos].setPosition(pos, PosInfo);
    }
};
CMathContent.prototype.Shift_Range = function(Dx, Dy, _CurLine, _CurRange, _CurPage)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    this.Bounds.ShiftPos(CurLine, CurRange, Dx, Dy);

	CParagraphContentWithParagraphLikeContent.prototype.Shift_Range.call(this, Dx, Dy, _CurLine, _CurRange, _CurPage);
};
CMathContent.prototype.SetParent = function(Parent, ParaMath)
{
    this.Parent   = Parent;
    this.ParaMath = ParaMath;
};
///// properties /////
CMathContent.prototype.hidePlaceholder = function(flag)
{
    this.plhHide = flag;
};
///////// RunPrp, CtrPrp

CMathContent.prototype.Get_FirstTextPr = function()
{
    return this.Content[0].Get_FirstTextPr();
};
CMathContent.prototype.getFirstRPrp  = function()
{
    return this.Content[0].Get_CompiledPr(true);
};
CMathContent.prototype.GetCtrPrp = function()       // for placeholder
{
    var ctrPrp = new CTextPr();
    if(!this.bRoot)
        ctrPrp.Merge( this.Parent.Get_CompiledCtrPrp_2() );

    return ctrPrp;
};
CMathContent.prototype.IsAccent = function()
{
    var result = false;

    if(!this.bRoot)
        result = this.Parent.IsAccent();

    return result;
};
////////////////////////
/// For Para Math
CMathContent.prototype.GetParent = function()
{
    return this.Parent ? this.Parent.GetParent() : null;
};
CMathContent.prototype.SetArgSize = function(val)
{
	History.Add(new CChangesMathContentArgSize(this, this.ArgSize.GetValue(), val));
	this.ArgSize.SetValue(val);
};
CMathContent.prototype.GetArgSize = function()
{
    return this.ArgSize.value;
};
/////////   Перемещение     ////////////

// Поиск позиции, селект
CMathContent.prototype.IsSelectedAll = function(Props)
{
    var bFirst = false, bEnd = false;

    if(this.Selection.StartPos == 0 && this.Selection.EndPos == this.Content.length - 1)
    {
        if(this.Content[this.Selection.StartPos].Type == para_Math_Run)
            bFirst = this.Content[this.Selection.StartPos].IsSelectedAll(Props);
        else
            bFirst = true;

        if(this.Content[this.Selection.EndPos].Type == para_Math_Run)
            bEnd = this.Content[this.Selection.EndPos].IsSelectedAll(Props);
        else
            bEnd = true;
    }

    return bFirst && bEnd;
};

///////////////////////

CMathContent.prototype.Get_Id = function()
{
    return this.GetId();
};
CMathContent.prototype.GetId = function()
{
    return this.Id;
};
CMathContent.prototype.private_CorrectContent = function()
{
    var len = this.Content.length;

    var EmptyRun = null;
    var RPr      = null;
    var CurrPos = 0;

    while(CurrPos < len)
    {
        var Current = this.Content[CurrPos];

        var bLeftRun  = CurrPos > 0 ? this.Content[CurrPos-1].Type == para_Math_Run : false,
            bRightRun = CurrPos < len - 1 ? this.Content[CurrPos + 1].Type === para_Math_Run : false;

        var bCurrComp       = Current.Type == para_Math_Composition,
            bCurrEmptyRun   = Current.Type == para_Math_Run && Current.Is_Empty();

        var bDeleteEmptyRun = bCurrEmptyRun && (bLeftRun || bRightRun);

        if(bCurrComp && !bLeftRun)
        {
            EmptyRun = new ParaRun(null, true);
            EmptyRun.Set_RFont_ForMathRun();

            RPr = Current.Get_CtrPrp(false);
            EmptyRun.Apply_Pr(RPr);
            this.Internal_Content_Add(CurrPos, EmptyRun);
            CurrPos += 2;

        }
        else if(bDeleteEmptyRun && false == Current.Is_CheckingNearestPos()) // если NearPosArray не нулевой длины, то это вызов происходит на InsertContent, не удаляем пустые Run
        {
            var tempPos = this.CurPos;
            this.Remove_FromContent(CurrPos, 1);

            if (tempPos === CurrPos)
            {
                if (bLeftRun)
                {
                    this.CurPos = CurrPos - 1;
                    this.Content[this.CurPos].MoveCursorToEndPos(false);
                }
                else
                {
                    this.Content[this.CurPos].MoveCursorToStartPos();
                }
            }
        }
        else
        {
            CurrPos++;
        }

        len = this.Content.length;
    }

    if(len > 1)
    {
        if(this.Content[len - 1].Type == para_Math_Composition)
        {
            EmptyRun = new ParaRun(null, true);
            EmptyRun.Set_RFont_ForMathRun();

            RPr = this.Content[len - 1].Get_CtrPrp(false);
            EmptyRun.Apply_Pr(RPr);
            this.Internal_Content_Add(CurrPos, EmptyRun);
        }
    }

};
CMathContent.prototype.Correct_Content = function(bInnerCorrection)
{
    if (true === bInnerCorrection)
    {
        for (var nPos = 0, nCount = this.Content.length; nPos < nCount; nPos++)
        {
            if (para_Math_Composition === this.Content[nPos].Type)
                this.Content[nPos].Correct_Content(true);
        }
    }

    this.private_CorrectContent();

    // Если в контенте ничего нет, тогда добавляем пустой ран
    if (this.Content.length < 1)
    {
        var NewMathRun = new ParaRun(null, true);
        NewMathRun.Set_RFont_ForMathRun();
        this.Add_ToContent(0, NewMathRun);
    }

	// Если единственный элемент данного контента ран и он пустой, заполняем его плейсхолдером
	if (1 === this.Content.length && para_Math_Run === this.Content[0].Type && true === this.Content[0].Is_Empty())
	{
		this.Content[0].AddMathPlaceholder();
	}

	// Возможна ситуация, когда у нас остались лишние плейсхолдеры (либо их слишком много, либо они вообще не нужны)
	if (true !== this.IsPlaceholder())
	{
		var isEmptyContent = true;
		for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
		{
			if (para_Math_Run === this.Content[nPos].Type)
			{
				this.Content[nPos].RemoveMathPlaceholder();
				if (false === this.Content[nPos].Is_Empty())
					isEmptyContent = false;
			}
			else
			{
				isEmptyContent = false;
			}
		}

		if (isEmptyContent)
			this.SetPlaceholder();
	}

	if (this.CurPos >= this.Content.length)
		this.CurPos = this.Content.length - 1;
    if (this.CurPos < 0)
    	this.CurPos = 0;
};
CMathContent.prototype.Correct_ContentPos = function(nDirection)
{
    var nCurPos = this.CurPos;

    if (nCurPos < 0)
    {
        this.CurPos = 0;
        this.Content[0].MoveCursorToStartPos();
    }
    else if (nCurPos > this.Content.length - 1)
    {
        this.CurPos = this.Content.length - 1;
        this.Content[this.CurPos].MoveCursorToEndPos();
    }
    else if (para_Math_Run !== this.Content[nCurPos].Type)
    {
        if (nDirection > 0)
        {
            this.CurPos = nCurPos + 1;
            this.Content[this.CurPos].MoveCursorToStartPos();
        }
        else
        {
            this.CurPos = nCurPos - 1;
            this.Content[this.CurPos].MoveCursorToEndPos();
        }
    }
};

/// функции для работы с курсором
CMathContent.prototype.Cursor_Is_Start = function()
{
    var result = false;

    if( !this.Is_Empty() )
    {
        if(this.CurPos == 0)
            result = this.Content[0].Cursor_Is_Start();
    }

    return result;
};
CMathContent.prototype.Cursor_Is_End = function()
{
    var result = false;

    if(!this.Is_Empty())
    {
        var len = this.Content.length - 1;
        if(this.CurPos == len)
        {
            result = this.Content[len].Cursor_Is_End();
        }
    }

    return result;
};
//////////////////////////////////////

/////////////////////////
//  Text Properties
///////////////

CMathContent.prototype.Get_TextPr = function(ContentPos, Depth)
{
    var pos = ContentPos.Get(Depth);

    var TextPr;

    if(true !== this.bRoot && this.IsPlaceholder())
        TextPr = this.Parent.Get_CtrPrp(true);
    else
        TextPr = this.Content[pos].Get_TextPr(ContentPos, Depth + 1);

    return TextPr;
};
CMathContent.prototype.GetDirectTextPr = function()
{
	let textPr = null;
	
	if (this.IsPlaceholder())
	{
		// ничего не делаем
	}
	else if (this.Selection.Use)
	{
		let startPos = this.Selection.StartPos;
		let endPos   = this.Selection.EndPos;
		if (startPos > endPos)
		{
			startPos = this.Selection.EndPos;
			endPos   = this.Selection.StartPos;
		}
		
		for (let pos = startPos; pos <= endPos; ++pos)
		{
			let curTextPr = this.Content[pos].GetDirectTextPr(false);
			if (!curTextPr)
				continue;
			
			if (!textPr)
				textPr = curTextPr.Copy();
			else
				textPr = textPr.Compare(curTextPr);
		}
		
		if (!textPr)
			textPr = this.Parent.Get_CtrPrp();
	}
	else
	{
		let pos = this.CurPos;
		
		if (0 <= pos && pos < this.Content.length)
			textPr = this.Content[pos].GetDirectTextPr();
	}
	
	if (this.Parent)
	{
		if (!textPr)
			textPr = this.Parent.Get_CtrPrp();
		else
			textPr = this.Parent.Get_CtrPrp().Copy().Compare(textPr);
	}
	
	return textPr;
};
CMathContent.prototype.Get_ParentCtrRunPr = function(bCopy)
{
    return this.Parent.Get_CtrPrp(bCopy);
};
CMathContent.prototype.Get_CompiledTextPr = function(Copy, bAll)
{
    var TextPr = null;

    if(true !== this.bRoot && this.IsPlaceholder())
    {
        TextPr = this.Parent.Get_CompiledCtrPrp_2();
    }
    else if (this.Selection.Use || bAll == true)
    {
        var StartPos, EndPos;

        if(bAll == true)
        {
            StartPos = 0;
            EndPos = this.Content.length - 1;
        }
        else
        {
            StartPos = this.Selection.StartPos;
            EndPos   = this.Selection.EndPos;

            if ( StartPos > EndPos )
            {
                StartPos = this.Selection.EndPos;
                EndPos   = this.Selection.StartPos;
            }
        }


        // пропускаем пустые рана только для случая, когда есть селект

        while ( null === TextPr && StartPos <= EndPos )
        {
            var bComp = this.Content[StartPos].Type == para_Math_Composition,
                bEmptyRun = this.Content[StartPos].Type == para_Math_Run && true === this.Content[StartPos].IsSelectionEmpty();

            if(bComp || !bEmptyRun || bAll)    //пропускаем пустые Run
                TextPr = this.Content[StartPos].Get_CompiledTextPr(true);

            StartPos++;
        }

        while(this.Content[EndPos].Type == para_Math_Run && true === this.Content[EndPos].IsSelectionEmpty() && StartPos < EndPos + 1 && bAll == false) //пропускаем пустые Run
        {
            EndPos--;
        }


        for ( var CurPos = StartPos; CurPos < EndPos + 1; CurPos++ )
        {
            //var CurTextPr = this.Content[CurPos].Get_CompiledPr(false);
            var CurTextPr = this.Content[CurPos].Get_CompiledTextPr(false);

            if ( null !== CurTextPr )
                TextPr = TextPr.Compare( CurTextPr );
        }
    }
    else
    {
        var CurPos = this.CurPos;

        if ( CurPos >= 0 && CurPos < this.Content.length )
            TextPr = this.Content[CurPos].Get_CompiledTextPr(Copy);
    }

    return TextPr;
};
CMathContent.prototype.GetMathTextPrForMenu = function(ContentPos, Depth)
{
    var pos = ContentPos.Get(Depth);

    return this.Content[pos].GetMathTextPrForMenu(ContentPos, Depth + 1);
};
CMathContent.prototype.Apply_TextPr = function(TextPr, IncFontSize, ApplyToAll, StartPos, EndPos)
{
    if ( true === ApplyToAll )
    {
        for ( var i = 0; i < this.Content.length; i++ )
            this.Content[i].Apply_TextPr( TextPr, IncFontSize, true );
    }
    else
    {
        var StartPos, EndPos, bMenu = false;

        if(StartPos !== undefined && EndPos !== undefined)
        {
            StartPos = StartPos;
            EndPos   = EndPos;

            bMenu = true;
        }
        else
        {
            StartPos = this.Selection.StartPos;
            EndPos   = this.Selection.EndPos;
        }

        var NewRuns;
        var LRun, CRun, RRun;

        var bSelectOneElement = this.Selection.Use && StartPos == EndPos;

        var FirstPos = this.Selection.Use ? Math.min(StartPos, EndPos) : this.CurPos;

        if(FirstPos == 0 && this.bRoot)
            this.ParaMath.SetRecalcCtrPrp(this.Content[0]);

        if( ( !this.Selection.Use && !bMenu ) || (bSelectOneElement && this.Content[StartPos].Type == para_Math_Run) ) // TextPr меняем только в одном Run
        {
            var Pos = !this.Selection.Use ? this.CurPos :  StartPos;

            NewRuns = this.Content[Pos].Apply_TextPr(TextPr, IncFontSize, false);

            LRun = NewRuns[0];
            CRun = NewRuns[1];
            RRun = NewRuns[2];

            var CRunPos = Pos;

            if(LRun !== null)
            {
                this.Internal_Content_Add(Pos+1, CRun);
                CRunPos = Pos + 1;
            }

            if(RRun !== null)
            {
                this.Internal_Content_Add(CRunPos+1, RRun);
            }

            this.CurPos             = CRunPos;
            this.Selection.StartPos = CRunPos;
            this.Selection.EndPos   = CRunPos;

        }
        else if(bSelectOneElement && this.Content[StartPos].Type == para_Math_Composition)  // заселекчен только один мат. объект
        {
            this.Content[StartPos].Apply_TextPr(TextPr, IncFontSize, true);
        }
        else
        {

            if(StartPos > EndPos)
            {
                var temp = StartPos;
                StartPos = EndPos;
                EndPos = temp;
            }


            for(var i = StartPos + 1; i < EndPos; i++)
                this.Content[i].Apply_TextPr(TextPr, IncFontSize, true );


            if(this.Content[EndPos].Type == para_Math_Run)
            {
                NewRuns = this.Content[EndPos].Apply_TextPr(TextPr, IncFontSize, false);

                // LRun - null
                CRun = NewRuns[1];
                RRun = NewRuns[2];

                if(RRun !== null)
                {
                    this.Internal_Content_Add(EndPos+1, RRun);
                }

            }
            else
                this.Content[EndPos].Apply_TextPr(TextPr, IncFontSize, true);


            if(this.Content[StartPos].Type == para_Math_Run)
            {
                NewRuns = this.Content[StartPos].Apply_TextPr(TextPr, IncFontSize, false);

                LRun = NewRuns[0];
                CRun = NewRuns[1];
                // RRun - null


                if(LRun !== null)
                {
                    this.Internal_Content_Add(StartPos+1, CRun);
                }

            }
            else
                this.Content[StartPos].Apply_TextPr(TextPr, IncFontSize, true);


            var bStartComposition = this.Content[StartPos].Type == para_Math_Composition || (this.Content[StartPos].Is_Empty() && this.Content[StartPos + 1].Type == para_Math_Composition);
            var bEndCompostion    = this.Content[EndPos].Type == para_Math_Composition || (this.Content[EndPos].Is_Empty()   && this.Content[EndPos - 1].Type == para_Math_Composition);

            if(!bStartComposition)
            {
                if(this.Selection.StartPos < this.Selection.EndPos && true === this.Content[this.Selection.StartPos].IsSelectionEmpty(true) )
                    this.Selection.StartPos++;
                else if (this.Selection.EndPos < this.Selection.StartPos && true === this.Content[this.Selection.EndPos].IsSelectionEmpty(true) )
                    this.Selection.EndPos++;
            }


            if(!bEndCompostion)
            {
                if(this.Selection.StartPos < this.Selection.EndPos && true === this.Content[this.Selection.EndPos].IsSelectionEmpty(true) )
                    this.Selection.EndPos--;
                else if (this.Selection.EndPos < this.Selection.StartPos && true === this.Content[this.Selection.StartPos].IsSelectionEmpty(true) )
                    this.Selection.StartPos--;
            }

        }
    }

};
CMathContent.prototype.Set_MathTextPr2 = function(TextPr, MathPr, bAll, StartPos, Count)
{
    if(bAll)
    {
        StartPos = 0;
        Count = this.Content.length - 1;
    }

    if(Count < 0 || StartPos + Count > this.Content.length - 1)
        return;

    for(var pos = StartPos; pos <= StartPos + Count; pos++)
        this.Content[pos].Set_MathTextPr2(TextPr, MathPr, true);

};
CMathContent.prototype.IsNormalTextInRuns = function()
{
    var flag = true;

    if(this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos,
            EndPos   = this.Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        for(var i = StartPos; i < EndPos+1; i++)
        {
            var curr = this.Content[i],
                currType = curr.Type;
            if(currType == para_Math_Composition || (currType == para_Math_Run && false == curr.IsNormalText()))
            {
                flag = false;
                break;
            }
        }
    }
    else
        flag = false;

    return flag;
};
CMathContent.prototype.Internal_Content_Add = function(Pos, Item, bUpdatePosition)
{
	Item.Set_ParaMath(this.ParaMath);

	if (Item.SetParagraph)
		Item.SetParagraph(this.Paragraph);

	Item.Parent = this;
	Item.Recalc_RunsCompiledPr();

	History.Add(new CChangesMathContentAddItem(this, Pos, [Item]));
	this.Content.splice(Pos, 0, Item);

	this.private_UpdatePosOnAdd(Pos, bUpdatePosition);
};
CMathContent.prototype.private_UpdatePosOnAdd = function(Pos, bUpdatePosition)
{
    if(bUpdatePosition !== false)
    {
        if ( this.CurPos >= Pos )
            this.CurPos++;

        if ( this.Selection.StartPos >= Pos )
            this.Selection.StartPos++;

        if ( this.Selection.EndPos >= Pos )
            this.Selection.EndPos++;

        this.private_CorrectSelectionPos();
        this.private_CorrectCurPos();
    }

    // Обновляем позиции в NearestPos
    var NearPosLen = this.NearPosArray.length;
    for ( var Index = 0; Index < NearPosLen; Index++ )
    {
        var HyperNearPos = this.NearPosArray[Index];
        var ContentPos = HyperNearPos.NearPos.ContentPos;
        var Depth      = HyperNearPos.Depth;

        if (ContentPos.Data[Depth] >= Pos)
            ContentPos.Data[Depth]++;
    }
};
CMathContent.prototype.private_CorrectSelectionPos = function()
{
    this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, this.Selection.StartPos));
    this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, this.Selection.EndPos));
};
CMathContent.prototype.private_CorrectCurPos = function()
{
    if (this.Content.length <= 0)
    {
        this.CurPos = 0;
        return;
    }

    if (this.CurPos > this.Content.length - 1)
    {
        this.CurPos = this.Content.length - 1;

        if (para_Math_Run === this.Content[this.CurPos].Type)
            this.Content[this.CurPos].MoveCursorToEndPos(false);
    }

    if (this.CurPos < 0)
    {
        this.CurPos = this.Content.length - 1;

        if (para_Math_Run === this.Content[this.CurPos].Type)
            this.Content[this.CurPos].MoveCursorToStartPos();
    }
};
CMathContent.prototype.Correct_ContentCurPos = function()
{
    this.private_CorrectCurPos();

    for(var Pos = 0; Pos < this.Content.length; Pos++)
    {
        if(this.Content[Pos].Type == para_Math_Composition)
            this.Content[Pos].Correct_ContentCurPos();
    }
};
CMathContent.prototype.SplitContent = function(NewContent, ContentPos, Depth)
{
    var Pos = ContentPos.Get(Depth);

    if(para_Math_Run === this.Content[Pos].Type)
    {
        var NewRun = this.Content[Pos].Split(ContentPos, Depth+1);
        NewContent.Add_ToContent(0, NewRun);

        var len = this.Content.length;
        if(Pos < len - 1)
        {
            NewContent.Concat_ToEnd( this.Content.slice(Pos + 1) );
            this.Remove_FromContent(Pos+1, len - Pos - 1);
        }
    }
    else
	{
		NewContent.Add_ToContent(0, new ParaRun(this.Paragraph, true));
	}
};
CMathContent.prototype.Add_ToContent = function(Pos, Item)
{
    if (Item && para_Run === Item.Type)
    {
        var MathRun = new ParaRun(Item.GetParagraph(), true);
        this.Internal_Content_Add(Pos, MathRun);
    }
    else
    {
        this.Internal_Content_Add(Pos, Item);
    }
};
CMathContent.prototype.Concat_ToEnd = function(NewItems)
{
    this.ConcatToContent(2, NewItems);
};
CMathContent.prototype.ConcatToContent = function(Pos, NewItems)
{
	if (NewItems != undefined && NewItems.length > 0)
	{
		var Count = NewItems.length;

		for (var i = 0; i < Count; i++)
		{
			NewItems[i].Set_ParaMath(this.ParaMath);
			NewItems[i].Parent = this;
			NewItems[i].Recalc_RunsCompiledPr();
		}

		History.Add(new CChangesMathContentAddItem(this, Pos, NewItems));

		var Array_start = this.Content.slice(0, Pos);
		var Array_end   = this.Content.slice(Pos);

		this.Content = Array_start.concat(NewItems, Array_end);
	}
};
CMathContent.prototype.Remove_FromContent = function(Pos, Count)
{
	if (Count <= 0)
		return;
	
	var DeletedItems = this.Content.splice(Pos, Count);
	History.Add(new CChangesMathContentRemoveItem(this, Pos, DeletedItems));

	// Обновим текущую позицию
	if (this.CurPos > Pos + Count)
		this.CurPos -= Count;
	else if (this.CurPos > Pos)
		this.CurPos = Pos;

	this.private_CorrectCurPos();
	this.private_UpdatePosOnRemove(Pos, Count);
};
CMathContent.prototype.Remove_Content = function(Pos, Count)
{
	this.Remove_FromContent(Pos, Count);
};
CMathContent.prototype.private_UpdatePosOnRemove = function(Pos, Count)
{
    // Обновим начало и конец селекта
    if (true === this.Selection.Use)
    {
        if (this.Selection.StartPos <= this.Selection.EndPos)
        {
            if (this.Selection.StartPos > Pos + Count)
                this.Selection.StartPos -= Count;
            else if (this.Selection.StartPos > Pos)
                this.Selection.StartPos = Pos;

            if (this.Selection.EndPos >= Pos + Count)
                this.Selection.EndPos -= Count;
            else if (this.Selection.EndPos >= Pos)
                this.Selection.EndPos = Math.max(0, Pos - 1);
        }
        else
        {
            if (this.Selection.StartPos >= Pos + Count)
                this.Selection.StartPos -= Count;
            else if (this.Selection.StartPos >= Pos)
                this.Selection.StartPos = Math.max(0, Pos - 1);

            if (this.Selection.EndPos > Pos + Count)
                this.Selection.EndPos -= Count;
            else if (this.Selection.EndPos > Pos)
                this.Selection.EndPos = Pos;
        }

        this.Selection.StartPos = Math.min(this.Content.length - 1, Math.max(0, this.Selection.StartPos));
        this.Selection.EndPos   = Math.min(this.Content.length - 1, Math.max(0, this.Selection.EndPos));
    }

    // Обновляем позиции в NearestPos
    var NearPosLen = this.NearPosArray.length;
    for (var Index = 0; Index < NearPosLen; Index++)
    {
        var HyperNearPos = this.NearPosArray[Index];
        var ContentPos = HyperNearPos.NearPos.ContentPos;
        var Depth      = HyperNearPos.Depth;

        if (ContentPos.Data[Depth] > Pos + Count)
            ContentPos.Data[Depth] -= Count;
        else if (ContentPos.Data[Depth] > Pos)
            ContentPos.Data[Depth] = Math.max(0 , Pos);
    }
};
CMathContent.prototype.Get_Default_TPrp = function()
{
    return this.ParaMath.Get_Default_TPrp();
};
/////////////////////////
CMathContent.prototype.Is_Empty = function()
{
    return this.Content.length == 0;
};
CMathContent.prototype.Copy = function(Selected, oPr)
{
    var NewContent = new CMathContent();
    this.CopyTo(NewContent, Selected, oPr);
    return NewContent;
};
CMathContent.prototype.CopyTo = function(OtherContent, Selected, oPr)
{
    var nStartPos, nEndPos;

    if(true === Selected)
    {
        if(this.Selection.StartPos < this.Selection.EndPos)
        {
            nStartPos = this.Selection.StartPos;
            nEndPos   = this.Selection.EndPos;
        }
        else
        {
            nStartPos = this.Selection.EndPos;
            nEndPos   = this.Selection.StartPos;
        }
    }
    else
    {
        nStartPos = 0;
        nEndPos   = this.Content.length - 1;
    }

    OtherContent.plHid = this.plhHide;
    OtherContent.SetArgSize(this.ArgSize.GetValue());

    for(var nPos = nStartPos; nPos <= nEndPos; nPos++)
    {
        var oElement;
        if(this.Content[nPos].Type == para_Math_Run)
            oElement = this.Content[nPos].Copy(Selected, oPr);
        else
            oElement = this.Content[nPos].Copy(false, oPr);

        OtherContent.Internal_Content_Add(OtherContent.Content.length, oElement);
    }
    if(oPr && oPr.Comparison)
    {
        if (oPr.SkipUpdateInfo) {
            oPr.Comparison.saveReviewInfo(OtherContent, OtherContent);
        } else if (oPr.bSaveCustomReviewType) {
            oPr.Comparison.saveCustomReviewInfo(OtherContent, OtherContent, oPr.Comparison.nInsertChangesType);
        } else {
            oPr.Comparison.updateReviewInfo(OtherContent, oPr.Comparison.nInsertChangesType);
        }
    }
};
CMathContent.prototype.getElem = function(nNum)
{
    return this.Content[nNum];
};
CMathContent.prototype.GetLastElement = function()
{
    var pos = this.Content.length - 1;

    while(this.Content[pos].Type == para_Math_Run && this.Content[pos].Is_Empty() && pos > 0)
    {
        pos--;
    }

    var last = this.Content[pos].Type == para_Math_Run ? this.Content[pos] : this.Content[pos].GetLastElement();

    return last;
};
CMathContent.prototype.GetFirstElement = function()
{
    var pos = 0;

    while(this.Content[pos].Type == para_Math_Run && this.Content[pos].Is_Empty() && pos < this.Content.length - 1)
    {
        pos++;
    }

    var first = this.Content[pos].Type == para_Math_Run ? this.Content[pos] : this.Content[pos].GetFirstElement();

    return first;
};
////////////////////////////////////////////////////////////////
CMathContent.prototype.Write_ToBinary2 = function(Writer)
{
    Writer.WriteLong(AscDFH.historyitem_type_MathContent);

    // Long : Id
    Writer.WriteString2(this.Id);
};
CMathContent.prototype.Read_FromBinary2 = function(Reader)
{
    // Long : Id
    this.Id = Reader.GetString2();
};
CMathContent.prototype.Refresh_RecalcData = function()
{
    if(this.ParaMath !== null)
        this.ParaMath.Refresh_RecalcData(); // Refresh_RecalcData сообщает родительскому классу, что у него произошли изменения, нужно пересчитать
};
CMathContent.prototype.InsertMathContent = function(oMathContent, nPos, isSelect)
{
	if (!oMathContent || !oMathContent.Content)
		return;
	
	if (!this.ParaMath || !this.ParaMath.Paragraph)
		isSelect = false;

	if (undefined === nPos)
		nPos = this.CurPos;
	
	let nCount = oMathContent.Content.length;
	for (let nIndex = 0; nIndex < nCount; ++nIndex)
	{
		let oElement = oMathContent.Content[nIndex];
		this.AddToContent(nPos + nIndex, oElement, true);

		if (isSelect)
			oElement.SelectAll(1);
		else
			oElement.RemoveSelection();
	}
	
	let newPos = nPos + nCount + 1;
	if (this.Content[newPos])
		this.Content[newPos].MoveCursorToStartPos();
	else
		this.Content[--newPos].MoveCursorToEndPos();
	
	if (isSelect)
	{
		this.Selection.Use      = true;
		this.Selection.StartPos = nPos;
		this.Selection.EndPos   = nPos + nCount - 1;
		this.CurPos             = newPos;

		if (!this.bRoot)
			this.ParentElement.Select_MathContent(this);
		else
			this.ParaMath.bSelectionUse = true;

		this.ParaMath.SelectThisElement(1, true);
	}
	else
	{
		this.ParaMath.SetThisElementCurrent();
		this.RemoveSelection();
		
		if (!this.bRoot)
			this.ParentElement.SetCurrentMathContent(this);
		
		this.CurPos = newPos;
	}
	this.Correct_Content(true);
	this.Correct_ContentPos(-1);
};
CMathContent.prototype.Set_ParaMath = function(ParaMath, Parent)
{
    this.Parent   = Parent;
    this.ParaMath = ParaMath;

    for (var Index = 0, Count = this.Content.length; Index < Count; Index++)
    {
        this.Content[Index].Set_ParaMath(ParaMath, this);
    }
};
CMathContent.prototype.Load_FromMenu = function(Type, Paragraph, TextPr, oSelectedContent)
{
    this.Paragraph = Paragraph;

    var Pr = {
    	ctrPrp : TextPr ? TextPr.Copy() : new CTextPr()
    };

    Pr.ctrPrp.Italic = true;
    Pr.ctrPrp.RFonts.SetAll("Cambria Math", -1);

    var MainType = Type >> 24;

    if (MainType === c_oAscMathMainType.Symbol)
		this.private_LoadFromMenuSymbol(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Fraction)
		this.private_LoadFromMenuFraction(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Script)
        this.private_LoadFromMenuScript(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Radical)
        this.private_LoadFromMenuRadical(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Integral)
        this.private_LoadFromMenuIntegral(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.LargeOperator)
        this.private_LoadFromMenuLargeOperator(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Bracket)
        this.private_LoadFromMenuBracket(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Function)
        this.private_LoadFromMenuFunction(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Accent)
        this.private_LoadFromMenuAccent(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.LimitLog)
        this.private_LoadFromMenuLimitLog(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Operator)
        this.private_LoadFromMenuOperator(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Matrix)
        this.private_LoadFromMenuMatrix(Type, Pr, oSelectedContent);
    else if(MainType == c_oAscMathMainType.Empty_Content)
		this.private_LoadFromMenuDefaultText(Type, Pr, oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuSymbol = function(Type, Pr)
{
    var Code = -1;

    switch (Type)
    {
        case c_oAscMathType.Symbol_pm            : Code = 0x00B1; break;
        case c_oAscMathType.Symbol_infinity      : Code = 0x221E; break;
        case c_oAscMathType.Symbol_equals        : Code = 0x003D; break;
        case c_oAscMathType.Symbol_neq           : Code = 0x2260; break;
        case c_oAscMathType.Symbol_about         : Code = 0x007E; break;
        case c_oAscMathType.Symbol_times         : Code = 0x00D7; break;
        case c_oAscMathType.Symbol_div           : Code = 0x00F7; break;
        case c_oAscMathType.Symbol_factorial     : Code = 0x0021; break;
        case c_oAscMathType.Symbol_propto        : Code = 0x221D; break;
        case c_oAscMathType.Symbol_less          : Code = 0x003C; break;
        case c_oAscMathType.Symbol_ll            : Code = 0x226A; break;
        case c_oAscMathType.Symbol_greater       : Code = 0x003E; break;
        case c_oAscMathType.Symbol_gg            : Code = 0x226B; break;
        case c_oAscMathType.Symbol_leq           : Code = 0x2264; break;
        case c_oAscMathType.Symbol_geq           : Code = 0x2265; break;
        case c_oAscMathType.Symbol_mp            : Code = 0x2213; break;
        case c_oAscMathType.Symbol_cong          : Code = 0x2245; break;
        case c_oAscMathType.Symbol_approx        : Code = 0x2248; break;
        case c_oAscMathType.Symbol_equiv         : Code = 0x2261; break;
        case c_oAscMathType.Symbol_forall        : Code = 0x2200; break;
        case c_oAscMathType.Symbol_additional    : Code = 0x2201; break;
        case c_oAscMathType.Symbol_partial       : Code = 0x1D715; break;
        case c_oAscMathType.Symbol_sqrt          : this.Add_Radical(Pr, null, null); break;
        case c_oAscMathType.Symbol_cbrt          : this.Add_Radical({ctrPrp : Pr.ctrPrp, type : DEGREE_RADICAL}, null, "3"); break;
        case c_oAscMathType.Symbol_qdrt          : this.Add_Radical({ctrPrp : Pr.ctrPrp, type : DEGREE_RADICAL}, null, "4"); break;
        case c_oAscMathType.Symbol_cup           : Code = 0x222A; break;
        case c_oAscMathType.Symbol_cap           : Code = 0x2229; break;
        case c_oAscMathType.Symbol_emptyset      : Code = 0x2205; break;
        case c_oAscMathType.Symbol_percent       : Code = 0x0025; break;
        case c_oAscMathType.Symbol_degree        : Code = 0x00B0; break;
        case c_oAscMathType.Symbol_fahrenheit    : Code = 0x2109; break;
        case c_oAscMathType.Symbol_celsius       : Code = 0x2103; break;
        case c_oAscMathType.Symbol_inc           : Code = 0x2206; break;
        case c_oAscMathType.Symbol_nabla         : Code = 0x2207; break;
        case c_oAscMathType.Symbol_exists        : Code = 0x2203; break;
        case c_oAscMathType.Symbol_notexists     : Code = 0x2204; break;
        case c_oAscMathType.Symbol_in            : Code = 0x2208; break;
        case c_oAscMathType.Symbol_ni            : Code = 0x220B; break;
        case c_oAscMathType.Symbol_leftarrow     : Code = 0x2190; break;
        case c_oAscMathType.Symbol_uparrow       : Code = 0x2191; break;
        case c_oAscMathType.Symbol_rightarrow    : Code = 0x2192; break;
        case c_oAscMathType.Symbol_downarrow     : Code = 0x2193; break;
        case c_oAscMathType.Symbol_leftrightarrow: Code = 0x2194; break;
        case c_oAscMathType.Symbol_therefore     : Code = 0x2234; break;
        case c_oAscMathType.Symbol_plus          : Code = 0x002B; break;
        case c_oAscMathType.Symbol_minus         : Code = 0x2212; break;
        case c_oAscMathType.Symbol_not           : Code = 0x00AC; break;
        case c_oAscMathType.Symbol_ast           : Code = 0x2217; break;
        case c_oAscMathType.Symbol_bullet        : Code = 0x2219; break;
        case c_oAscMathType.Symbol_vdots         : Code = 0x22EE; break;
        case c_oAscMathType.Symbol_cdots         : Code = 0x22EF; break;
        case c_oAscMathType.Symbol_rddots        : Code = 0x22F0; break;
        case c_oAscMathType.Symbol_ddots         : Code = 0x22F1; break;
        case c_oAscMathType.Symbol_aleph         : Code = 0x2135; break;
        case c_oAscMathType.Symbol_beth          : Code = 0x2136; break;
        case c_oAscMathType.Symbol_QED           : Code = 0x220E; break;
        case c_oAscMathType.Symbol_alpha         : Code = 0x03B1; break;
        case c_oAscMathType.Symbol_beta          : Code = 0x03B2; break;
        case c_oAscMathType.Symbol_gamma         : Code = 0x03B3; break;
        case c_oAscMathType.Symbol_delta         : Code = 0x03B4; break;
        case c_oAscMathType.Symbol_varepsilon    : Code = 0x03B5; break;
        case c_oAscMathType.Symbol_epsilon       : Code = 0x03F5; break;
        case c_oAscMathType.Symbol_zeta          : Code = 0x03B6; break;
        case c_oAscMathType.Symbol_eta           : Code = 0x03B7; break;
        case c_oAscMathType.Symbol_theta         : Code = 0x03B8; break;
        case c_oAscMathType.Symbol_vartheta      : Code = 0x03D1; break;
        case c_oAscMathType.Symbol_iota          : Code = 0x03B9; break;
        case c_oAscMathType.Symbol_kappa         : Code = 0x03BA; break;
        case c_oAscMathType.Symbol_lambda        : Code = 0x03BB; break;
        case c_oAscMathType.Symbol_mu            : Code = 0x03BC; break;
        case c_oAscMathType.Symbol_nu            : Code = 0x03BD; break;
        case c_oAscMathType.Symbol_xsi           : Code = 0x03BE; break;
        case c_oAscMathType.Symbol_o             : Code = 0x03BF; break;
        case c_oAscMathType.Symbol_pi            : Code = 0x03C0; break;
        case c_oAscMathType.Symbol_varpi         : Code = 0x03D6; break;
        case c_oAscMathType.Symbol_rho           : Code = 0x03C1; break;
        case c_oAscMathType.Symbol_varrho        : Code = 0x03F1; break;
        case c_oAscMathType.Symbol_sigma         : Code = 0x03C3; break;
        case c_oAscMathType.Symbol_varsigma      : Code = 0x03C2; break;
        case c_oAscMathType.Symbol_tau           : Code = 0x03C4; break;
        case c_oAscMathType.Symbol_upsilon       : Code = 0x03C5; break;
        case c_oAscMathType.Symbol_varphi        : Code = 0x03C6; break;
        case c_oAscMathType.Symbol_phi           : Code = 0x03D5; break;
        case c_oAscMathType.Symbol_chi           : Code = 0x03C7; break;
        case c_oAscMathType.Symbol_psi           : Code = 0x03C8; break;
        case c_oAscMathType.Symbol_omega         : Code = 0x03C9; break;
        case c_oAscMathType.Symbol_Alpha         : Code = 0x0391; break;
        case c_oAscMathType.Symbol_Beta          : Code = 0x0392; break;
        case c_oAscMathType.Symbol_Gamma         : Code = 0x0393; break;
        case c_oAscMathType.Symbol_Delta         : Code = 0x0394; break;
        case c_oAscMathType.Symbol_Epsilon       : Code = 0x0395; break;
        case c_oAscMathType.Symbol_Zeta          : Code = 0x0396; break;
        case c_oAscMathType.Symbol_Eta           : Code = 0x0397; break;
        case c_oAscMathType.Symbol_Theta         : Code = 0x0398; break;
        case c_oAscMathType.Symbol_Iota          : Code = 0x0399; break;
        case c_oAscMathType.Symbol_Kappa         : Code = 0x039A; break;
        case c_oAscMathType.Symbol_Lambda        : Code = 0x039B; break;
        case c_oAscMathType.Symbol_Mu            : Code = 0x039C; break;
        case c_oAscMathType.Symbol_Nu            : Code = 0x039D; break;
        case c_oAscMathType.Symbol_Xsi           : Code = 0x039E; break;
        case c_oAscMathType.Symbol_O             : Code = 0x039F; break;
        case c_oAscMathType.Symbol_Pi            : Code = 0x03A0; break;
        case c_oAscMathType.Symbol_Rho           : Code = 0x03A1; break;
        case c_oAscMathType.Symbol_Sigma         : Code = 0x03A3; break;
        case c_oAscMathType.Symbol_Tau           : Code = 0x03A4; break;
        case c_oAscMathType.Symbol_Upsilon       : Code = 0x03A5; break;
        case c_oAscMathType.Symbol_Phi           : Code = 0x03A6; break;
        case c_oAscMathType.Symbol_Chi           : Code = 0x03A7; break;
        case c_oAscMathType.Symbol_Psi           : Code = 0x03A8; break;
        case c_oAscMathType.Symbol_Omega         : Code = 0x03A9; break;
    }

    if (-1 !== Code)
    {
        var TextPr, MathPr;

        if (this.Content.length <= 0)
			this.Correct_Content();

        if(this.Content.length > 0 && this.Content[this.CurPos].Type == para_Math_Run && this.IsSelectionEmpty() == true) // находимся в Run, селект отсутствует
        {
            TextPr = this.Content[this.CurPos].Get_TextPr();
            TextPr.RFonts.SetAll("Cambria Math", -1);          //  на данный момент добавляются символы исключительно из Cambria Math
            MathPr = this.Content[this.CurPos].GetMathPr();
        }

        this.Add_Symbol(Code, TextPr, MathPr);
    }
};
CMathContent.prototype.private_LoadFromMenuFraction = function(Type, Pr, oSelectedContent)
{
	var oFraction = null;
    switch (Type)
    {
        case c_oAscMathType.FractionVertical   : oFraction = this.Add_Fraction(Pr, null, null); break;
        case c_oAscMathType.FractionDiagonal   : oFraction = this.Add_Fraction({ctrPrp : Pr.ctrPrp, type : SKEWED_FRACTION}, null, null); break;
        case c_oAscMathType.FractionHorizontal : oFraction = this.Add_Fraction({ctrPrp : Pr.ctrPrp, type : LINEAR_FRACTION}, null, null); break;
        case c_oAscMathType.FractionSmall:
            var oBox = new CBox(Pr);
            this.Add_Element(oBox);
            var BoxMathContent = oBox.getBase();
            BoxMathContent.SetArgSize(-1);
            oFraction = BoxMathContent.Add_Fraction(Pr, null, null);
            break;

        case c_oAscMathType.FractionDifferential_1: this.Add_Fraction(Pr, "dx", "dy"); break;
        case c_oAscMathType.FractionDifferential_2: this.Add_Fraction(Pr, String.fromCharCode(916) + "y", String.fromCharCode(916) + "x"); break;
        case c_oAscMathType.FractionDifferential_3: this.Add_Fraction(Pr, String.fromCharCode(8706) + "y", String.fromCharCode(8706) + "x"); break;
        case c_oAscMathType.FractionDifferential_4: this.Add_Fraction(Pr, String.fromCharCode(948) + "y", String.fromCharCode(948) + "x"); break;
        case c_oAscMathType.FractionPi_2          : this.Add_Fraction(Pr, String.fromCharCode(960), "2"); break;
    }

    if (oFraction && oSelectedContent)
		oFraction.getNumeratorMathContent().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuScript = function(Type, Pr, oSelectedContent)
{
	var oScript = null;
    switch (Type)
    {
        case c_oAscMathType.ScriptSup: oScript = this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, null, null, null); break;
        case c_oAscMathType.ScriptSub: oScript = this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, null, null, null); break;
        case c_oAscMathType.ScriptSubSup: oScript = this.Add_Script(true, {ctrPrp : Pr.ctrPrp, type : DEGREE_SubSup}, null, null, null); break;
        case c_oAscMathType.ScriptSubSupLeft: oScript = this.Add_Script(true, {ctrPrp : Pr.ctrPrp, type : DEGREE_PreSubSup}, null, null, null); break;
        case c_oAscMathType.ScriptCustom_1:
            Pr.type = DEGREE_SUBSCRIPT;
            var Script = this.Add_Script(false, Pr, "x", null, null);
            var SubMathContent = Script.getLowerIterator();
            Pr.type = DEGREE_SUPERSCRIPT;
            SubMathContent.Add_Script(false, Pr, "y", "2", null);
            break;

        case c_oAscMathType.ScriptCustom_2: this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "e", "-i" + String.fromCharCode(969) + "t", null); break;
        case c_oAscMathType.ScriptCustom_3: this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "x", "2", null); break;
        case c_oAscMathType.ScriptCustom_4: this.Add_Script(true, {ctrPrp : Pr.ctrPrp, type : DEGREE_PreSubSup}, "Y", "n", "1"); break;
    }

	if (oScript && oSelectedContent)
		oScript.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuRadical = function(Type, Pr, oSelectedContent)
{
	var oRadical = null;
    switch (Type)
    {
        case c_oAscMathType.RadicalSqrt:
            Pr.type    = SQUARE_RADICAL;
            Pr.degHide = true;
            oRadical = this.Add_Radical(Pr, null, null);
            break;

        case c_oAscMathType.RadicalRoot_n:
            Pr.type = DEGREE_RADICAL;
			oRadical = this.Add_Radical(Pr, null, null);
            break;

        case c_oAscMathType.RadicalRoot_2:
            Pr.type = DEGREE_RADICAL;
			oRadical = this.Add_Radical(Pr, null, "2");
            break;

        case c_oAscMathType.RadicalRoot_3:
            Pr.type = DEGREE_RADICAL;
			oRadical = this.Add_Radical(Pr, null, "3");
            break;

        case c_oAscMathType.RadicalCustom_1:
            var Fraction = this.Add_Fraction(Pr, null, null);
            var NumMathContent = Fraction.getNumeratorMathContent();
            var DenMathContent = Fraction.getDenominatorMathContent();

            NumMathContent.Add_Text("-b" + String.fromCharCode(177), this.Paragraph);
            Pr.type    = SQUARE_RADICAL;
            Pr.degHide = true;
            var Radical = NumMathContent.Add_Radical(Pr, null, null);
            var RadicalBaseMathContent = Radical.getBase();
            RadicalBaseMathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "b", "2", null);
            RadicalBaseMathContent.Add_Text("-4ac", this.Paragraph);

            DenMathContent.Add_Text("2a", this.Paragraph);

            break;
        case c_oAscMathType.RadicalCustom_2:
            Pr.type    = SQUARE_RADICAL;
            Pr.degHide = true;
            var Radical = this.Add_Radical(Pr, null, null);
            var BaseMathContent = Radical.getBase();

            var ScriptPr = {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT};
            BaseMathContent.Add_Script(false, ScriptPr, "a", "2", null);
            BaseMathContent.Add_Text("+", this.Paragraph);
            BaseMathContent.Add_Script(false, ScriptPr, "b", "2", null);
            break;
    }

	if (oRadical && oSelectedContent)
		oRadical.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuIntegral = function(Type, Pr, oSelectedContent)
{
	var oIntegral = null;
    switch(Type)
    {
        case c_oAscMathType.Integral:                           oIntegral = this.Add_Integral(1, false, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralSubSup:                     oIntegral = this.Add_Integral(1, false, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralCenterSubSup:               oIntegral = this.Add_Integral(1, false, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralDouble:                     oIntegral = this.Add_Integral(2, false, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralDoubleSubSup:               oIntegral = this.Add_Integral(2, false, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralDoubleCenterSubSup:         oIntegral = this.Add_Integral(2, false, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralTriple:                     oIntegral = this.Add_Integral(3, false, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralTripleSubSup:               oIntegral = this.Add_Integral(3, false, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralTripleCenterSubSup:         oIntegral = this.Add_Integral(3, false, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOriented:                   oIntegral = this.Add_Integral(1,  true, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedSubSup:             oIntegral = this.Add_Integral(1,  true, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedCenterSubSup:       oIntegral = this.Add_Integral(1,  true, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedDouble:             oIntegral = this.Add_Integral(2,  true, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedDoubleSubSup:       oIntegral = this.Add_Integral(2,  true, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedDoubleCenterSubSup: oIntegral = this.Add_Integral(2,  true, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedTriple:             oIntegral = this.Add_Integral(3,  true, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedTripleSubSup:       oIntegral = this.Add_Integral(3,  true, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedTripleCenterSubSup: oIntegral = this.Add_Integral(3,  true, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.Integral_dx:     Pr.diff = 1; this.Add_Box(Pr, "dx"); break;
        case c_oAscMathType.Integral_dy:     Pr.diff = 1; this.Add_Box(Pr, "dy"); break;
        case c_oAscMathType.Integral_dtheta: Pr.diff = 1; this.Add_Box(Pr, "d" + String.fromCharCode(952)); break;
    }

	if (oIntegral && oSelectedContent)
		oIntegral.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuLargeOperator = function(Type, Pr, oSelectedContent)
{
	var oOperator = null;
    switch(Type)
    {
        case c_oAscMathType.LargeOperator_Sum:              oOperator = this.Add_LargeOperator(1, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_CenterSubSup: oOperator = this.Add_LargeOperator(1, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_SubSup:       oOperator = this.Add_LargeOperator(1, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_CenterSub:    oOperator = this.Add_LargeOperator(1, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_Sub:          oOperator = this.Add_LargeOperator(1, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Prod:              oOperator = this.Add_LargeOperator(2, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_CenterSubSup: oOperator = this.Add_LargeOperator(2, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_SubSup:       oOperator = this.Add_LargeOperator(2, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_CenterSub:    oOperator = this.Add_LargeOperator(2, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_Sub:          oOperator = this.Add_LargeOperator(2, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_CoProd:              oOperator = this.Add_LargeOperator(3, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_CenterSubSup: oOperator = this.Add_LargeOperator(3, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_SubSup:       oOperator = this.Add_LargeOperator(3, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_CenterSub:    oOperator = this.Add_LargeOperator(3, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_Sub:          oOperator = this.Add_LargeOperator(3, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Union:              oOperator = this.Add_LargeOperator(4, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_CenterSubSup: oOperator = this.Add_LargeOperator(4, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_SubSup:       oOperator = this.Add_LargeOperator(4, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_CenterSub:    oOperator = this.Add_LargeOperator(4, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_Sub:          oOperator = this.Add_LargeOperator(4, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Intersection:              oOperator = this.Add_LargeOperator(5, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_CenterSubSup: oOperator = this.Add_LargeOperator(5, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_SubSup:       oOperator = this.Add_LargeOperator(5, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_CenterSub:    oOperator = this.Add_LargeOperator(5, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_Sub:          oOperator = this.Add_LargeOperator(5, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Disjunction:              oOperator = this.Add_LargeOperator(6, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_CenterSubSup: oOperator = this.Add_LargeOperator(6, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_SubSup:       oOperator = this.Add_LargeOperator(6, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_CenterSub:    oOperator = this.Add_LargeOperator(6, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_Sub:          oOperator = this.Add_LargeOperator(6, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Conjunction:              oOperator = this.Add_LargeOperator(7, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_CenterSubSup: oOperator = this.Add_LargeOperator(7, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_SubSup:       oOperator = this.Add_LargeOperator(7, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_CenterSub:    oOperator = this.Add_LargeOperator(7, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_Sub:          oOperator = this.Add_LargeOperator(7, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Custom_1:
            var Sum = this.Add_LargeOperator(1, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, "k");
            var BaseMathContent = Sum.getBaseMathContent();
            var Delimiter = BaseMathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, [null]);
            var DelimiterMathContent = Delimiter.getElementMathContent(0);
            DelimiterMathContent.Add_Fraction({ctrPrp: Pr.ctrPrp, type : NO_BAR_FRACTION}, "n", "k");
            break;

        case c_oAscMathType.LargeOperator_Custom_2:
            this.Add_LargeOperator(1, NARY_UndOvr, false, false, Pr.ctrPrp, null, "n", "i=0");
            break;

        case c_oAscMathType.LargeOperator_Custom_3:
            var Sum = this.Add_LargeOperator(1, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null);
            var SubMathContent = Sum.getSubMathContent();
            SubMathContent.Add_EqArray({ctrPrp: Pr.ctrPrp, row : 2}, 2, ["0≤ i ≤ m", "0< j < n"]);
            var BaseMathContent = Sum.getBaseMathContent();
            BaseMathContent.Add_Text("P", this.Paragraph);
            BaseMathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, ["i, j"]);
            break;

        case c_oAscMathType.LargeOperator_Custom_4:
            var Prod = this.Add_LargeOperator(2, NARY_UndOvr, false, false, Pr.ctrPrp, null, "n", "k=1");
            var BaseMathContent = Prod.getBaseMathContent();
            BaseMathContent.Add_Script(false, {ctrPrp: Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, "A", null, "k");
            break;

        case c_oAscMathType.LargeOperator_Custom_5:
            var Union = this.Add_LargeOperator(4, NARY_UndOvr, false, false, Pr.ctrPrp, null, "m", "n=1");
            var BaseMathContent = Union.getBaseMathContent();
            var Delimiter = BaseMathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, [null]);
            BaseMathContent = Delimiter.getElementMathContent(0);
            BaseMathContent.Add_Script(false, {ctrPrp: Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, "X", null, "n");
            BaseMathContent.Add_Text(String.fromCharCode(8745), this.Paragraph);
            BaseMathContent.Add_Script(false, {ctrPrp: Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, "Y", null, "n");
            break;
    }

	if (oOperator && oSelectedContent)
		oOperator.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuBracket = function(Type, Pr, oSelectedContent)
{
	var oBracket  = null;
	var oFraction = null;

    switch(Type)
    {
        case c_oAscMathType.Bracket_Round                 : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,  null); break;
        case c_oAscMathType.Bracket_Square                : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    91,    93); break;
        case c_oAscMathType.Bracket_Curve                 : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,   125); break;
        case c_oAscMathType.Bracket_Angle                 : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10216, 10217); break;
        case c_oAscMathType.Bracket_LowLim                : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x230A, 0x230B); break;
        case c_oAscMathType.Bracket_UppLim                : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x2308, 0x2309); break;
        case c_oAscMathType.Bracket_Line                  : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   124,   124); break;
        case c_oAscMathType.Bracket_LineDouble            : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  8214,  8214); break;
        case c_oAscMathType.Bracket_Square_OpenOpen       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    91,    91); break;
        case c_oAscMathType.Bracket_Square_CloseClose     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    93,    93); break;
        case c_oAscMathType.Bracket_Square_CloseOpen      : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    93,    91); break;
        case c_oAscMathType.Bracket_SquareDouble          : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10214, 10215); break;

        case c_oAscMathType.Bracket_Round_Delimiter_2     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 2, [null, null], null, null); break;
        case c_oAscMathType.Bracket_Curve_Delimiter_2     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 2, [null, null], 123, 125); break;
        case c_oAscMathType.Bracket_Angle_Delimiter_2     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 2, [null, null], 10216, 10217); break;
        case c_oAscMathType.Bracket_Angle_Delimiter_3     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 3, [null, null, null], 10216, 10217); break;

        case c_oAscMathType.Bracket_Round_OpenNone        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,    -1); break;
        case c_oAscMathType.Bracket_Round_NoneOpen        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  null); break;
        case c_oAscMathType.Bracket_Square_OpenNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    91,    -1); break;
        case c_oAscMathType.Bracket_Square_NoneOpen       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,    93); break;
        case c_oAscMathType.Bracket_Curve_OpenNone        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1); break;
        case c_oAscMathType.Bracket_Curve_NoneOpen        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,   125); break;
        case c_oAscMathType.Bracket_Angle_OpenNone        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10216,    -1); break;
        case c_oAscMathType.Bracket_Angle_NoneOpen        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1, 10217); break;
        case c_oAscMathType.Bracket_LowLim_OpenNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x230A,    -1); break;
        case c_oAscMathType.Bracket_LowLim_NoneNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  0x230B); break;
        case c_oAscMathType.Bracket_UppLim_OpenNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x2308,    -1); break;
        case c_oAscMathType.Bracket_UppLim_NoneOpen       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  0x2309); break;
        case c_oAscMathType.Bracket_Line_OpenNone         : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   124,    -1); break;
        case c_oAscMathType.Bracket_Line_NoneOpen         : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,   124); break;
        case c_oAscMathType.Bracket_LineDouble_OpenNone   : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  8214,    -1); break;
        case c_oAscMathType.Bracket_LineDouble_NoneOpen   : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  8214); break;
        case c_oAscMathType.Bracket_SquareDouble_OpenNone : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10214,    -1); break;
        case c_oAscMathType.Bracket_SquareDouble_NoneOpen : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1, 10215); break;

        case c_oAscMathType.Bracket_Custom_1:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oBracket = BaseMathContent.Add_EqArray({ctrPrp : Pr.ctrPrp, row : 2}, 2, [null, null]);
            break;

        case c_oAscMathType.Bracket_Custom_2:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oBracket = BaseMathContent.Add_EqArray({ctrPrp : Pr.ctrPrp, row : 3}, 3, [null, null, null]);
            break;

        case c_oAscMathType.Bracket_Custom_3:
			oFraction = this.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, null, null);
            break;

        case c_oAscMathType.Bracket_Custom_4:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,  null);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oFraction = BaseMathContent.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, null, null);
            break;

        case c_oAscMathType.Bracket_Custom_5:
            this.Add_Text("f", this.Paragraph);
            this.Add_DelimiterEx(Pr.ctrPrp, 1, ["x"],  null,  null);
            this.Add_Text("=", this.Paragraph);
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oBracket = BaseMathContent.Add_EqArray({ctrPrp : Pr.ctrPrp, row : 2}, 2, ["-x,  &x<0", "x,  &x" + String.fromCharCode(8805) + "0"]);
            break;

        case c_oAscMathType.Bracket_Custom_6:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,  null);
            var BaseMathContent = Delimiter.getElementMathContent(0);
            BaseMathContent.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, "n", "k");
            break;

        case c_oAscMathType.Bracket_Custom_7:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  10216,  10217);
            var BaseMathContent = Delimiter.getElementMathContent(0);
            BaseMathContent.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, "n", "k");
            break;
    }

	if (oBracket && oSelectedContent)
		oBracket.getElementMathContent(0).private_FillSelectedContent(oSelectedContent);
    else if (oFraction && oSelectedContent)
		oFraction.getNumeratorMathContent().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuFunction = function(Type, Pr, oSelectedContent)
{
	var oFunction = null;
    switch(Type)
    {
        case c_oAscMathType.Function_Sin : oFunction = this.Add_Function(Pr, "sin", null); break;
        case c_oAscMathType.Function_Cos : oFunction = this.Add_Function(Pr, "cos", null); break;
        case c_oAscMathType.Function_Tan : oFunction = this.Add_Function(Pr, "tan", null); break;
        case c_oAscMathType.Function_Csc : oFunction = this.Add_Function(Pr, "csc", null); break;
        case c_oAscMathType.Function_Sec : oFunction = this.Add_Function(Pr, "sec", null); break;
        case c_oAscMathType.Function_Cot : oFunction = this.Add_Function(Pr, "cot", null); break;

        case c_oAscMathType.Function_1_Sin : oFunction = this.Add_Function_1(Pr, "sin", null); break;
        case c_oAscMathType.Function_1_Cos : oFunction = this.Add_Function_1(Pr, "cos", null); break;
        case c_oAscMathType.Function_1_Tan : oFunction = this.Add_Function_1(Pr, "tan", null); break;
        case c_oAscMathType.Function_1_Csc : oFunction = this.Add_Function_1(Pr, "csc", null); break;
        case c_oAscMathType.Function_1_Sec : oFunction = this.Add_Function_1(Pr, "sec", null); break;
        case c_oAscMathType.Function_1_Cot : oFunction = this.Add_Function_1(Pr, "cot", null); break;

        case c_oAscMathType.Function_Sinh : oFunction = this.Add_Function(Pr, "sinh", null); break;
        case c_oAscMathType.Function_Cosh : oFunction = this.Add_Function(Pr, "cosh", null); break;
        case c_oAscMathType.Function_Tanh : oFunction = this.Add_Function(Pr, "tanh", null); break;
        case c_oAscMathType.Function_Csch : oFunction = this.Add_Function(Pr, "csch", null); break;
        case c_oAscMathType.Function_Sech : oFunction = this.Add_Function(Pr, "sech", null); break;
        case c_oAscMathType.Function_Coth : oFunction = this.Add_Function(Pr, "coth", null); break;

        case c_oAscMathType.Function_1_Sinh : oFunction = this.Add_Function_1(Pr, "sinh", null); break;
        case c_oAscMathType.Function_1_Cosh : oFunction = this.Add_Function_1(Pr, "cosh", null); break;
        case c_oAscMathType.Function_1_Tanh : oFunction = this.Add_Function_1(Pr, "tanh", null); break;
        case c_oAscMathType.Function_1_Csch : oFunction = this.Add_Function_1(Pr, "csch", null); break;
        case c_oAscMathType.Function_1_Sech : oFunction = this.Add_Function_1(Pr, "sech", null); break;
        case c_oAscMathType.Function_1_Coth : oFunction = this.Add_Function_1(Pr, "coth", null); break;

        case c_oAscMathType.Function_Custom_1 : this.Add_Function(Pr, "sin", String.fromCharCode(952)); break;
        case c_oAscMathType.Function_Custom_2 : this.Add_Function(Pr, "cos", "2x"); break;
        case c_oAscMathType.Function_Custom_3 :
            var Theta = String.fromCharCode(952);
            this.Add_Function(Pr, "tan", Theta);
            this.Add_Text("=", this.Paragraph);
            var Fraction = this.Add_Fraction(Pr, null, null);
            var NumMathContent = Fraction.getNumeratorMathContent();
            var DenMathContent = Fraction.getDenominatorMathContent();
            NumMathContent.Add_Function(Pr, "sin", Theta);
            DenMathContent.Add_Function(Pr, "cos", Theta);
            break;
    }

	if (oFunction && oSelectedContent)
		oFunction.getArgument().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuAccent = function(Type, Pr, oSelectedContent)
{
	var oAccent = null;
    switch(Type)
    {
        case c_oAscMathType.Accent_Dot       : oAccent = this.Add_Accent(Pr.ctrPrp, 775, null); break;
        case c_oAscMathType.Accent_DDot      : oAccent = this.Add_Accent(Pr.ctrPrp, 776, null); break;
        case c_oAscMathType.Accent_DDDot     : oAccent = this.Add_Accent(Pr.ctrPrp, 8411, null); break;
        case c_oAscMathType.Accent_Hat       : oAccent = this.Add_Accent(Pr.ctrPrp, null, null); break;
        case c_oAscMathType.Accent_Check     : oAccent = this.Add_Accent(Pr.ctrPrp, 780, null); break;
        case c_oAscMathType.Accent_Accent    : oAccent = this.Add_Accent(Pr.ctrPrp, 769, null); break;
        case c_oAscMathType.Accent_Grave     : oAccent = this.Add_Accent(Pr.ctrPrp, 768, null); break;
        case c_oAscMathType.Accent_Smile     : oAccent = this.Add_Accent(Pr.ctrPrp, 774, null); break;
        case c_oAscMathType.Accent_Tilde     : oAccent = this.Add_Accent(Pr.ctrPrp, 771, null); break;
        case c_oAscMathType.Accent_Bar       : oAccent = this.Add_Accent(Pr.ctrPrp, 773, null); break;
        case c_oAscMathType.Accent_DoubleBar : oAccent = this.Add_Accent(Pr.ctrPrp, 831, null); break;

        case c_oAscMathType.Accent_CurveBracketTop : oAccent = this.Add_GroupCharacter({ctrPrp : Pr.ctrPrp, chr : 9182, pos : VJUST_TOP, vertJc : VJUST_BOT}, null ); break;
        case c_oAscMathType.Accent_CurveBracketBot : oAccent = this.Add_GroupCharacter({ctrPrp : Pr.ctrPrp}, null ); break;
        case c_oAscMathType.Accent_GroupTop:
            var Limit = this.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_UP}, null, null);
            var MathContent = Limit.getFName();
			oAccent = MathContent.Add_GroupCharacter({ctrPrp : Pr.ctrPrp, chr : 9182, pos : VJUST_TOP, vertJc : VJUST_BOT}, null );
            break;

        case c_oAscMathType.Accent_GroupBot:
            var Limit = this.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_LOW}, null, null);
            var MathContent = Limit.getFName();
			oAccent = MathContent.Add_GroupCharacter({ctrPrp : Pr.ctrPrp}, null );
            break;

        case c_oAscMathType.Accent_ArrowL  : oAccent = this.Add_Accent(Pr.ctrPrp, 8406, null); break;
        case c_oAscMathType.Accent_ArrowR  : oAccent = this.Add_Accent(Pr.ctrPrp, 8407, null); break;
        case c_oAscMathType.Accent_ArrowD  : oAccent = this.Add_Accent(Pr.ctrPrp, 8417, null); break;
        case c_oAscMathType.Accent_HarpoonL: oAccent = this.Add_Accent(Pr.ctrPrp, 8400, null); break;
        case c_oAscMathType.Accent_HarpoonR: oAccent = this.Add_Accent(Pr.ctrPrp, 8401, null); break;

        case c_oAscMathType.Accent_BorderBox :
			oAccent = this.Add_BorderBox(Pr, null);
            break;

        case c_oAscMathType.Accent_BorderBoxCustom :
            var BorderBox = this.Add_BorderBox(Pr, null);
            var MathContent = BorderBox.getBase();
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "a", "2", null);
            MathContent.Add_Text("=", this.Paragraph);
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "b", "2", null);
            MathContent.Add_Text("+", this.Paragraph);
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "c", "2", null);
            break;

        case c_oAscMathType.Accent_BarTop : oAccent = this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, null); break;
        case c_oAscMathType.Accent_BarBot : oAccent = this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_BOT}, null); break;

        case c_oAscMathType.Accent_Custom_1 :
            this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, "A");
            break;

        case c_oAscMathType.Accent_Custom_2 :
            this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, "ABC");
            break;

        case c_oAscMathType.Accent_Custom_3 :
            this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, "x" + String.fromCharCode(8853) + "y");
            break;
    }

	if (oAccent && oSelectedContent)
		oAccent.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuLimitLog = function(Type, Pr, oSelectedContent)
{
	var oFunction = null;
    switch(Type)
    {
        case c_oAscMathType.LimitLog_LogBase:
			oFunction = this.Add_Function(Pr, null, null);
            var MathContent = oFunction.getFName();
            var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, null, null, null);
            MathContent = Script.getBase();
            MathContent.Add_Text("log", this.Paragraph, STY_PLAIN);
            break;

        case c_oAscMathType.LimitLog_Log: oFunction = this.Add_Function(Pr, "log", null); break;
        case c_oAscMathType.LimitLog_Lim: oFunction = this.Add_FunctionWithLimit(Pr, "lim", null, null); break;
        case c_oAscMathType.LimitLog_Min: oFunction = this.Add_FunctionWithLimit(Pr, "min", null, null); break;
        case c_oAscMathType.LimitLog_Max: oFunction = this.Add_FunctionWithLimit(Pr, "max", null, null); break;
        case c_oAscMathType.LimitLog_Ln : oFunction = this.Add_Function(Pr, "ln", null); break;

        case c_oAscMathType.LimitLog_Custom_1:
            var Function = this.Add_FunctionWithLimit(Pr, "lim", "n" + String.fromCharCode(8594,8734), null);
            var MathContent = Function.getArgument();
            var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, null, "n", null);
            MathContent = Script.getBase();
            var Delimiter = MathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, [null]);
            MathContent = Delimiter.getElementMathContent(0);
            MathContent.Add_Text("1+", this.Paragraph);
            MathContent.Add_Fraction({ctrPrp : Pr.ctrPrp}, "1", "n");
            break;

        case c_oAscMathType.LimitLog_Custom_2:
            var Function = this.Add_FunctionWithLimit(Pr, "max", "0" + String.fromCharCode(8804) + "x" + String.fromCharCode(8804) + "1", null);
            var MathContent = Function.getArgument();
            MathContent.Add_Text("x", this.Paragraph);
            var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "e", null, null);
            MathContent = Script.getUpperIterator();
            MathContent.Add_Text("-", this.Paragraph);
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "x", "2", null);
            break;
    }

	if (oFunction && oSelectedContent)
		oFunction.getArgument().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuOperator = function(Type, Pr, oSelectedContent)
{
	var oAccent = null;
    switch(Type)
    {
        case c_oAscMathType.Operator_ColonEquals     : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(0x2254)); break;
        case c_oAscMathType.Operator_EqualsEquals    : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, "=="); break;
        case c_oAscMathType.Operator_PlusEquals      : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, "+="); break;
        case c_oAscMathType.Operator_MinusEquals     : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, "-="); break;
        case c_oAscMathType.Operator_Definition      : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(8797)); break;
        case c_oAscMathType.Operator_UnitOfMeasure   : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(8798)); break;
        case c_oAscMathType.Operator_DeltaEquals     : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(8796)); break;
        case c_oAscMathType.Operator_ArrowL_Top      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8592, null); break;
        case c_oAscMathType.Operator_ArrowR_Top      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8594, null); break;
        case c_oAscMathType.Operator_ArrowL_Bot      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8592, null); break;
        case c_oAscMathType.Operator_ArrowR_Bot      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8594, null); break;
        case c_oAscMathType.Operator_DoubleArrowL_Top: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8656, null); break;
        case c_oAscMathType.Operator_DoubleArrowR_Top: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8658, null); break;
        case c_oAscMathType.Operator_DoubleArrowL_Bot: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8656, null); break;
        case c_oAscMathType.Operator_DoubleArrowR_Bot: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8658, null); break;
        case c_oAscMathType.Operator_ArrowD_Top      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8596, null); break;
        case c_oAscMathType.Operator_ArrowD_Bot      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8596, null); break;
        case c_oAscMathType.Operator_DoubleArrowD_Top: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8660, null); break;
        case c_oAscMathType.Operator_DoubleArrowD_Bot: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8660, null); break;
        case c_oAscMathType.Operator_Custom_1        : this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8594, "yields"); break;
        case c_oAscMathType.Operator_Custom_2        : this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8594, String.fromCharCode(8710)); break;
    }

	if (oAccent && oSelectedContent)
		oAccent.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuMatrix = function(Type, Pr, oSelectedContent)
{
	var oMatrix = null;
    switch(Type)
    {
        case c_oAscMathType.Matrix_1_2: oMatrix = this.Add_Matrix(Pr.ctrPrp, 1, 2, false, []); break;
        case c_oAscMathType.Matrix_2_1: oMatrix = this.Add_Matrix(Pr.ctrPrp, 2, 1, false, []); break;
        case c_oAscMathType.Matrix_1_3: oMatrix = this.Add_Matrix(Pr.ctrPrp, 1, 3, false, []); break;
        case c_oAscMathType.Matrix_3_1: oMatrix = this.Add_Matrix(Pr.ctrPrp, 3, 1, false, []); break;
        case c_oAscMathType.Matrix_2_2: oMatrix = this.Add_Matrix(Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_3: oMatrix = this.Add_Matrix(Pr.ctrPrp, 2, 3, false, []); break;
        case c_oAscMathType.Matrix_3_2: oMatrix = this.Add_Matrix(Pr.ctrPrp, 3, 2, false, []); break;
        case c_oAscMathType.Matrix_3_3: oMatrix = this.Add_Matrix(Pr.ctrPrp, 3, 3, false, []); break;

        case c_oAscMathType.Matrix_Dots_Center   : this.Add_Text(String.fromCharCode(8943), this.Paragraph); break;
        case c_oAscMathType.Matrix_Dots_Baseline : this.Add_Text(String.fromCharCode(8230), this.Paragraph); break;
        case c_oAscMathType.Matrix_Dots_Vertical : this.Add_Text(String.fromCharCode(8942), this.Paragraph); break;
        case c_oAscMathType.Matrix_Dots_Diagonal : this.Add_Text(String.fromCharCode(8945), this.Paragraph); break;

        case c_oAscMathType.Matrix_Identity_2         : this.Add_Matrix(Pr.ctrPrp, 2, 2, false, ["1", "0", "0", "1"]); break;
        case c_oAscMathType.Matrix_Identity_2_NoZeros : this.Add_Matrix(Pr.ctrPrp, 2, 2, true, ["1", null, null, "1"]); break;
        case c_oAscMathType.Matrix_Identity_3         : this.Add_Matrix(Pr.ctrPrp, 3, 3, false, ["1", "0", "0", "0", "1", "0", "0", "0", "1"]); break;
        case c_oAscMathType.Matrix_Identity_3_NoZeros : this.Add_Matrix(Pr.ctrPrp, 3, 3, true, ["1", null, null, null, "1", null, null, null, "1"]); break;

        case c_oAscMathType.Matrix_2_2_RoundBracket  : oMatrix = this.Add_MatrixWithBrackets(null, null, Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_2_SquareBracket : oMatrix = this.Add_MatrixWithBrackets(  91,   93, Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_2_LineBracket   : oMatrix = this.Add_MatrixWithBrackets( 124,  124, Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_2_DLineBracket  : oMatrix = this.Add_MatrixWithBrackets(8214, 8214, Pr.ctrPrp, 2, 2, false, []); break;

        case c_oAscMathType.Matrix_Flat_Round  : this.Add_MatrixWithBrackets(null, null, Pr.ctrPrp, 3, 3, false, [null, String.fromCharCode(8943), null, String.fromCharCode(8942), String.fromCharCode(8945), String.fromCharCode(8942), null, String.fromCharCode(8943), null]); break;
        case c_oAscMathType.Matrix_Flat_Square : this.Add_MatrixWithBrackets(  91,   93, Pr.ctrPrp, 3, 3, false, [null, String.fromCharCode(8943), null, String.fromCharCode(8942), String.fromCharCode(8945), String.fromCharCode(8942), null, String.fromCharCode(8943), null]); break;
    }

    if (oMatrix)
		oMatrix.getContentElement(0, 0).private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuDefaultText = function(Type, Pr, oSelectedContent)
{
	if (oSelectedContent)
		this.private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_FillSelectedContent = function(oSelectedContent)
{
	if (oSelectedContent instanceof ParaMath)
	{
		oSelectedContent.Root.CopyTo(this, false);
	}
	else if (oSelectedContent)
	{
		this.Add_Text(oSelectedContent, this.Paragraph);
	}
};
CMathContent.prototype.Add_Element = function(Element)
{
    this.Internal_Content_Add(this.CurPos, Element, false);
    this.CurPos++;
};
CMathContent.prototype.Add_Text = function(text, paragraph, mathStyle)
{
	if (!text)
		return;
	
	if (this.IsAddTextInLastParaRun(mathStyle))
	{
		this.Add_ToPrevParaRun(text);
		return;
	}
	
	var oMathRun = new AscWord.CRun(undefined, true);
	
	AscWord.TextToMathRunElements(text, function(item)
	{
		oMathRun.Add(item, true);
	});

	oMathRun.Set_RFont_ForMathRun();
	
	if (mathStyle)
		oMathRun.Math_Apply_Style(mathStyle);

	this.AddToContent(this.CurPos, oMathRun, false);
	this.CurPos++;
};
CMathContent.prototype.Add_ToPrevParaRun = function(text)
{
	let run = this.Content[this.Content.length - 1];

	AscWord.TextToMathRunElements(text, function(item)
	{
		run.Add(item, true);
	});
}
CMathContent.prototype.Add_TextOnPos = function(nPos, sText, MathStyle)
{
	if (!sText)
		return;

	let oMathRun = new AscWord.CRun(undefined, true);

	AscWord.TextToMathRunElements(sText, function(item)
	{
		oMathRun.Add(item, true);
	});

	oMathRun.Set_RFont_ForMathRun();

	if (undefined !== MathStyle && null !== MathStyle)
		oMathRun.Math_Apply_Style(MathStyle);

	this.Internal_Content_Add(nPos, oMathRun, false);
	this.CurPos++;

};
CMathContent.prototype.IsAddTextInLastParaRun = function(oMathStyle)
{
	if (this.Content.length > 0)
	{
		let oLastContent = this.Content[this.Content.length - 1];
		if (
			oLastContent instanceof ParaRun &&
			!oLastContent.IsPlaceholder() &&
			!oLastContent.Is_Empty() &&
			(oMathStyle === oLastContent.MathPrp.sty || oMathStyle === undefined)
		)
			return true;
	}
	return false
}
CMathContent.prototype.Add_Symbol = function(Code, TextPr, MathPr)
{
    var MathRun = new ParaRun(this.Paragraph, true);

    var Symbol = new CMathText(false);
    Symbol.add(Code);
    MathRun.Add(Symbol, true);

    if(TextPr !== undefined)
        MathRun.Apply_Pr(TextPr);

    if(MathPr !== undefined)
        MathRun.Set_MathPr(MathPr);

    this.Internal_Content_Add(this.CurPos, MathRun, false);
    this.CurPos++;
};
CMathContent.prototype.Add_Fraction = function(Pr, NumText, DenText)
{
    var Fraction = new CFraction(Pr);
    this.Add_Element(Fraction);

    var DenMathContent = Fraction.getDenominatorMathContent();
    DenMathContent.Add_Text(DenText, this.Paragraph);

    var NumMathContent = Fraction.getNumeratorMathContent();
    NumMathContent.Add_Text(NumText, this.Paragraph);

    return Fraction;
};
CMathContent.prototype.Add_Script = function(bSubSup, Pr, BaseText, SupText, SubText)
{
    var Script = null;

    if (bSubSup)
        Script = new CDegreeSubSup(Pr);
    else
        Script = new CDegree(Pr);

    this.Add_Element(Script);

    var MathContent = Script.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = Script.getUpperIterator();
    MathContent.Add_Text(SupText, this.Paragraph);

    MathContent = Script.getLowerIterator();
    MathContent.Add_Text(SubText, this.Paragraph);

    return Script;
};
CMathContent.prototype.Add_Radical = function(Pr, BaseText, DegreeText)
{
    var Radical = new CRadical(Pr);
    this.Add_Element(Radical);

    var MathContent = Radical.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = Radical.getDegree();
    MathContent.Add_Text(DegreeText, this.Paragraph);

    return Radical;
};
CMathContent.prototype.Add_NAry = function(Pr, BaseText, SupText, SubText)
{
    var NAry = new CNary(Pr);
    this.Add_Element(NAry);

    var MathContent = NAry.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = NAry.getSubMathContent();
    MathContent.Add_Text(SubText, this.Paragraph);

    MathContent = NAry.getSupMathContent();
    MathContent.Add_Text(SupText, this.Paragraph);

    return NAry;
};
CMathContent.prototype.Add_Integral = function(Dim, bOriented, limLoc, supHide, subHide, ctrPr, BaseText, SupText, SubText)
{
    var Pr = {ctrPrp : ctrPr};

    if (null !== limLoc)
        Pr.limLoc = limLoc;

    if (null !== supHide)
        Pr.supHide = supHide;

    if (null !== subHide)
        Pr.subHide = subHide;

    var chr = null;
    switch(Dim)
    {
        case 3: chr = (bOriented ? 8752 : 8749); break;
        case 2: chr = (bOriented ? 8751 : 8748); break;
        default:
        case 1: chr = (bOriented ? 8750 : null); break;
    }

    if (null !== chr)
        Pr.chr = chr;

    return this.Add_NAry(Pr, BaseText, SupText, SubText);
};
CMathContent.prototype.Add_LargeOperator = function(Type, limLoc, supHide, subHide, ctrPr, BaseText, SupText, SubText)
{
    var Pr = {ctrPrp : ctrPr};

    if (null !== limLoc)
        Pr.limLoc = limLoc;

    if (null !== supHide)
        Pr.supHide = supHide;

    if (null !== subHide)
        Pr.subHide = subHide;

    var chr = null;
    switch(Type)
    {
        default:
        case 1: chr = 8721; break;
        case 2: chr = 8719; break;
        case 3: chr = 8720; break;
        case 4: chr = 8899; break;
        case 5: chr = 8898; break;
        case 6: chr = 8897; break;
        case 7: chr = 8896; break;
    }

    if (null !== chr)
        Pr.chr = chr;

    return this.Add_NAry(Pr, BaseText, SupText, SubText);
};
CMathContent.prototype.Add_Delimiter = function(Pr, Count, aText)
{
    var Del = new CDelimiter(Pr);

    this.Add_Element(Del);

    for (var Index = 0; Index < Count; Index++)
    {
        var MathContent = Del.getElementMathContent(Index);
        MathContent.Add_Text(aText[Index], this.Paragraph);
    }

    return Del;
};
CMathContent.prototype.Add_DelimiterEx = function(ctrPr, Count, aText, begChr, endChr)
{
    var Pr =
    {
        ctrPrp : ctrPr,
        column : Count,
        begChr : begChr,
        endChr : endChr
    };

    return this.Add_Delimiter(Pr, Count, aText);
};
CMathContent.prototype.Add_EqArray = function(Pr, Count, aText)
{
    var EqArray = new CEqArray(Pr);

    this.Add_Element(EqArray);

    for (var Index = 0; Index < Count; Index++)
    {
        var MathContent = EqArray.getElementMathContent(Index);
        MathContent.Add_Text(aText[Index], this.Paragraph);
    }

    return EqArray;
};
CMathContent.prototype.Add_Box = function(Pr, BaseText)
{
    var Box = new CBox(Pr);
    this.Add_Element(Box);

    var MathContent = Box.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Box;
};
CMathContent.prototype.Add_BoxWithGroupChar = function(BoxPr, GroupPos, GroupChr, BaseText)
{
    var Box = this.Add_Box(BoxPr, null);
    var MathContent = Box.getBase();

    var oGroup = null;
    if (GroupPos === VJUST_TOP)
		oGroup = MathContent.Add_GroupCharacter({ctrPrp : BoxPr.ctrPrp, pos : GroupPos, chr : GroupChr}, BaseText);
    else
		oGroup = MathContent.Add_GroupCharacter({ctrPrp : BoxPr.ctrPrp, vertJc : GroupPos, chr : GroupChr}, BaseText);

    return oGroup;
};
CMathContent.prototype.Add_BorderBox = function(Pr, BaseText)
{
    var Box = new CBorderBox(Pr);
    this.Add_Element(Box);

    var MathContent = Box.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Box;
};
CMathContent.prototype.Add_Bar = function(Pr, BaseText)
{
    var Bar = new CBar(Pr);
    this.Add_Element(Bar);

    var MathContent = Bar.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Bar;
};
CMathContent.prototype.Add_Function = function(Pr, FName, BaseText)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    var MathContent = MathFunc.getFName();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_Function_1 = function(Pr, FName, BaseText)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    MathFunc.SetParagraph(this.Paragraph);

    var MathContent = MathFunc.getFName();
    var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, null, "-1", null);
    MathContent = Script.getBase();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_FunctionWithLimit = function(Pr, FName, LimitText, BaseText)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    MathFunc.SetParagraph(this.Paragraph);

    var MathContent = MathFunc.getFName();
    var Limit = MathContent.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_LOW}, null, LimitText);
    MathContent = Limit.getFName();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_FunctionWithTypeLimit = function(Pr, FName, LimitText, BaseText, LIMIT_TYPE)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    MathFunc.SetParagraph(this.Paragraph);

    var MathContent = MathFunc.getFName();
    var Limit = MathContent.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_TYPE}, null, LimitText);
    MathContent = Limit.getFName();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_Accent = function(ctrPr, chr, BaseText)
{
    var Pr =
    {
        ctrPrp : ctrPr,
        chr    : chr
    };
    var Accent = new CAccent(Pr);
    this.Add_Element(Accent);

    var MathContent = Accent.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Accent;
};
CMathContent.prototype.Add_GroupCharacter = function(Pr, BaseText)
{
    var Group = new CGroupCharacter(Pr);
    this.Add_Element(Group);

    var MathContent = Group.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Group;
};
CMathContent.prototype.Add_Limit = function(Pr, BaseText, LimitText)
{
    var Limit = new CLimit(Pr);
    this.Add_Element(Limit);

    var MathContent = Limit.getFName();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = Limit.getIterator();
    MathContent.Add_Text(LimitText, this.Paragraph);

    return Limit;
};
CMathContent.prototype.Add_Matrix = function(ctrPr, RowsCount, ColsCount, plcHide, aText)
{
    var Pr =
    {
        ctrPrp  : ctrPr,
        row     : RowsCount,
        mcs     : [{count : ColsCount, mcJc : MCJC_CENTER}],
        plcHide : plcHide
    };

    var Matrix = new CMathMatrix(Pr);
    this.Add_Element(Matrix);

    for (var RowIndex = 0; RowIndex < RowsCount; RowIndex++)
    {
        for (var ColIndex = 0; ColIndex < ColsCount; ColIndex++)
        {
            var MathContent = Matrix.getContentElement(RowIndex, ColIndex);
            MathContent.Add_Text(aText[RowIndex * ColsCount + ColIndex], this.Paragraph);
        }
    }

    return Matrix;
};
CMathContent.prototype.Add_MatrixWithBrackets = function(begChr, endChr, ctrPr, RowsCount, ColsCount, plcHide, aText)
{
    var Delimiter = this.Add_DelimiterEx(ctrPr, 1, [null], begChr, endChr);
    var MathContent = Delimiter.getElementMathContent(0);
    return MathContent.Add_Matrix(ctrPr, RowsCount, ColsCount, plcHide, aText);
};
CMathContent.prototype.Recalculate_CurPos = function(_X, _Y, CurrentRun, _CurRange, _CurLine, _CurPage, UpdateCurPos, UpdateTarget, ReturnTarget)
{
	if (-1 === this.StartLine)
	{
		return {
			X         : 0,
			Y         : 0,
			Height    : 0,
			PageNum   : 0,
			Internal  : {Line : 0, Page : 0, Range : 0},
			Transform : null
		};
	}

    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var _EndPos = ( true === CurrentRun ? Math.min( EndPos, this.CurPos ) : EndPos );

    return this.Content[_EndPos].Recalculate_CurPos(_X, _Y, CurrentRun, _CurRange, _CurLine, _CurPage, UpdateCurPos, UpdateTarget, ReturnTarget);
};
CMathContent.prototype.GetCurrentParaPos = function()
{
	if (this.CurPos >= 0 && this.CurPos < this.Content.length)
		return this.Content[this.CurPos].GetCurrentParaPos();

	return new CParaPos(this.StartRange, this.StartLine, 0, 0);
};
CMathContent.prototype.Get_ParaContentPos = function(bSelection, bStart, ContentPos, bUseCorrection)
{
    if (true === bUseCorrection && true === bSelection)
	{
		var nPos = false !== bStart ? this.Selection.StartPos : this.Selection.EndPos;

		if (para_Math_Run !== this.Content[nPos].Type
			&& (this.Selection.StartPos !== this.Selection.EndPos || true !== this.Content[nPos].Is_InnerSelection())
			&& ((true === bStart && nPos > 0) || (true !== bStart && nPos < this.Content.length - 1)))
		{
			if (true === bStart && nPos > 0)
			{
				ContentPos.Add(nPos - 1);
				this.Content[nPos - 1].Get_EndPos(false, ContentPos, ContentPos.GetDepth() + 1);
			}
			else
			{
				ContentPos.Add(nPos + 1);
				this.Content[nPos + 1].Get_StartPos(ContentPos, ContentPos.GetDepth() + 1);
			}
		}
		else
		{
			ContentPos.Add(nPos);

			if (undefined !== this.Content[nPos])
				this.Content[nPos].Get_ParaContentPos(bSelection, bStart, ContentPos, bUseCorrection);
		}
	}
	else
	{
		var nPos = (true !== bSelection ? this.CurPos : (false !== bStart ? this.Selection.StartPos : this.Selection.EndPos));
		nPos = Math.max(0, Math.min(nPos, this.Content.length - 1));

		ContentPos.Add(nPos);

		if (undefined !== this.Content[nPos])
			this.Content[nPos].Get_ParaContentPos(bSelection, bStart, ContentPos, bUseCorrection);
	}
};
CMathContent.prototype.Set_ParaContentPos = function(ContentPos, Depth)
{
    var CurPos = ContentPos.Get(Depth);

    // Делаем такие проверки, потому что после совместного редактирования, позиция может остаться старой, а
    // контент измениться.
    if (undefined === CurPos || CurPos < 0)
    {
        this.CurPos = 0;
        if (this.Content[this.CurPos])
        	this.Content[this.CurPos].MoveCursorToStartPos();
    }
    else if (CurPos > this.Content.length - 1)
    {
        this.CurPos = this.Content.length - 1;
        if (this.Content[this.CurPos])
        	this.Content[this.CurPos].MoveCursorToEndPos(false);
    }
    else
    {
        this.CurPos = CurPos;
        if (this.Content[this.CurPos])
        	this.Content[this.CurPos].Set_ParaContentPos(ContentPos, Depth + 1);
    }
};
CMathContent.prototype.IsSelectionEmpty = function()
{
    if (true !== this.Selection.Use)
        return true;

    if (this.Selection.StartPos === this.Selection.EndPos)
        return this.Content[this.Selection.StartPos].IsSelectionEmpty();

    return false;
};
CMathContent.prototype.GetSelectContent = function(isAll)
{
	if (true === isAll)
	{
		return {Content : this, Start : 0, End : this.Content.length - 1};
	}
    else if (false === this.Selection.Use)
    {
        if (this.Content.length - 1 >= this.CurPos && para_Math_Composition === this.Content[this.CurPos].Type)
            return this.Content[this.CurPos].GetSelectContent();
        else
            return {Content : this, Start : this.CurPos, End : this.CurPos};
    }
    else
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;

        if (StartPos > EndPos)
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        if (StartPos === EndPos && para_Math_Composition === this.Content[StartPos].Type && true === this.Content[StartPos].Is_InnerSelection())
            return this.Content[StartPos].GetSelectContent();

        return {Content : this, Start : StartPos, End : EndPos};
    }
};
CMathContent.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
	if ((this.IsPlaceholder() && UseContentPos)
		|| true !== this.ParentElement.Is_ContentUse(this))
		return false;

    if (false === UseContentPos && para_Math_Run === this.Content[this.Content.length - 1].Type)
    {
        // При переходе в новый контент встаем в его конец
        var CurPos = this.Content.length - 1;
        this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found = true;

        return true;
    }

    var CurPos = UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1;

    var bStepStart = false;
    if (CurPos > 0 || !this.Content[0].Cursor_Is_Start())
        bStepStart = true;

    this.Content[CurPos].Get_LeftPos(SearchPos, ContentPos, Depth + 1, UseContentPos);
    SearchPos.Pos.Update(CurPos, Depth);

    if (true === SearchPos.Found)
        return true;

    CurPos--;

    if (true === UseContentPos && para_Math_Composition === this.Content[CurPos + 1].Type)
    {
        // При выходе из формулы встаем в конец рана
        this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found = true;
        return true;
    }

    while (CurPos >= 0)
    {
        this.Content[CurPos].Get_LeftPos(SearchPos, ContentPos, Depth + 1, false);
        SearchPos.Pos.Update( CurPos, Depth );

        if (true === SearchPos.Found)
            return true;

        CurPos--;
    }

    if (true === bStepStart)
    {
        // Перед выходом из контента встаем в его начало
        this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(0, Depth);
        SearchPos.Found = true;

        return true;
    }

    return false;
};
CMathContent.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
	if ((this.IsPlaceholder() && UseContentPos)
		|| true !== this.ParentElement.Is_ContentUse(this))
		return false;

    if (false === UseContentPos && para_Math_Run === this.Content[0].Type)
    {
        // При переходе в новый контент встаем в его начало
        this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(0, Depth);
        SearchPos.Found = true;

        return true;
    }

    var CurPos = true === UseContentPos ? ContentPos.Get(Depth) : 0;

    var Count = this.Content.length;
    var bStepEnd = false;
    if (CurPos < Count - 1 || !this.Content[Count - 1].Cursor_Is_End())
        bStepEnd = true;

    this.Content[CurPos].Get_RightPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepEnd);
    SearchPos.Pos.Update( CurPos, Depth );

    if (true === SearchPos.Found)
        return true;

    CurPos++;

    if (true === UseContentPos && para_Math_Composition === this.Content[CurPos - 1].Type)
    {
        // При выходе из формулы встаем в начало рана
        this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found = true;
        return true;
    }

    while (CurPos < Count)
    {
        this.Content[CurPos].Get_RightPos(SearchPos, ContentPos, Depth + 1, false, StepEnd);
        SearchPos.Pos.Update(CurPos, Depth);

        if (true === SearchPos.Found)
            return true;

        CurPos++;
    }

    if (true === bStepEnd)
    {
        // Перед выходом из контента встаем в его конец
        this.Content[Count - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(Count - 1, Depth);
        SearchPos.Found = true;

        return true;
    }

    return false;
};
CMathContent.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepStart)
{
	if (true !== this.ParentElement.Is_ContentUse(this))
		return false;

	if (false === UseContentPos && para_Math_Run === this.Content[this.Content.length - 1].Type)
	{
		// При переходе в новый контент встаем в его конец
		var CurPos = this.Content.length - 1;
		this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(CurPos, Depth);
		SearchPos.Found		= true;
		SearchPos.UpdatePos = true;
		return true;
	}

	var CurPos = true === UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1;

	var bStepStart = false;
	if (CurPos > 0 || !this.Content[0].Cursor_Is_Start())
		bStepStart = true;

	this.Content[CurPos].Get_WordStartPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepStart);

	if (true === SearchPos.UpdatePos)
		SearchPos.Pos.Update( CurPos, Depth );

	if (true === SearchPos.Found)
		return;

	CurPos--;

	var bStepStartRun = false;
	if (true === UseContentPos && para_Math_Composition === this.Content[CurPos + 1].Type)
	{
		// При выходе из формулы встаем в конец рана
		this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(CurPos, Depth);
		SearchPos.Found		= true;
		SearchPos.UpdatePos	= true;
		return true;
	}
	else if (para_Math_Run === this.Content[CurPos + 1].Type && true === SearchPos.Shift)
		bStepStartRun = true;



	while (CurPos >= 0)
	{
		if (true !== bStepStartRun || para_Math_Run === this.Content[CurPos].Type)
		{
			var OldUpdatePos = SearchPos.UpdatePos;
			if (this.Content[CurPos].Type !== 49)
			{
				this.Content[CurPos - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
				SearchPos.Pos.Update(CurPos - 1, Depth);
				SearchPos.Found		= true;
				SearchPos.UpdatePos	= true;
				return;
			}
			else
			{
				this.Content[CurPos].Get_WordStartPos(SearchPos, ContentPos, Depth + 1, false, StepStart);

				if (true === SearchPos.UpdatePos)
					SearchPos.Pos.Update(CurPos, Depth);
				else
					SearchPos.UpdatePos = OldUpdatePos;

				if (true === SearchPos.Found)
					return;

				if (true === SearchPos.Shift)
					bStepStartRun = true;
			}
		}
		else
		{
			// Встаем в начало рана перед формулой
			this.Content[CurPos + 1].Get_StartPos(SearchPos.Pos, Depth + 1);
			SearchPos.Pos.Update(CurPos + 1, Depth);
			SearchPos.Found		= true;
			SearchPos.UpdatePos	= true;
			return true;
		}
		CurPos--;
	}

	if (true === bStepStart)
	{
		// Перед выходом из контента встаем в его начало
		this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(0, Depth);
		SearchPos.Found		= true;
		SearchPos.UpdatePos	= true;
		return true;
	}
};
CMathContent.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
	if (true !== this.ParentElement.Is_ContentUse(this))
		return false;

	if (false === UseContentPos && para_Math_Run === this.Content[0].Type)
	{
		// При переходе в новый контент встаем в его начало
		this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(0, Depth);
		SearchPos.Found		= true;
		SearchPos.UpdatePos	= true;
		return true;
	}

	var CurPos = true === UseContentPos ? ContentPos.Get(Depth) : 0;

	var Count = this.Content.length;
	var bStepEnd = false;
	if (CurPos < Count - 1 || !this.Content[Count - 1].Cursor_Is_End())
		bStepEnd = true;

	this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepEnd);

	if (true === SearchPos.UpdatePos)
		SearchPos.Pos.Update( CurPos, Depth);

	if (true === SearchPos.Found)
		return;

	CurPos++;

	var bStepEndRun = false;
	if (true === UseContentPos && para_Math_Composition === this.Content[CurPos - 1].Type)
	{
		// При выходе из формулы встаем в начало рана
		this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(CurPos, Depth);
		SearchPos.Found		= true;
		SearchPos.UpdatePos	= true;
		return true;
	}
	else if (para_Math_Run === this.Content[CurPos - 1].Type && true === SearchPos.Shift)
		bStepEndRun = true;

	while (CurPos < Count)
	{
		if (true !== bStepEndRun || para_Math_Run === this.Content[CurPos].Type)
		{
			var OldUpdatePos = SearchPos.UpdatePos;

			if (this.Content[CurPos].Type !== 49)
			{
				//this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, false, false);
				SearchPos.Pos.Update(CurPos+1, Depth);
				SearchPos.Pos.Update(0, Depth + 1);
				SearchPos.Found		= true;
				SearchPos.UpdatePos	= true;
				return;
			}
			else
			{
				this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, false, StepEnd);

				if (true === SearchPos.UpdatePos)
					SearchPos.Pos.Update(CurPos, Depth);
				else
					SearchPos.UpdatePos = OldUpdatePos;

				if (true === SearchPos.Found)
					return;

				if (true === SearchPos.Shift)
					bStepEndRun = true;
			}
		}
		else
		{
			// Встаем в конец рана перед формулой
			this.Content[CurPos - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
			SearchPos.Pos.Update(CurPos - 1, Depth);
			SearchPos.Found		= true;
			SearchPos.UpdatePos	= true;
			return true;
		}

		CurPos++;
	}

	if (true === bStepEnd)
	{
		// Перед выходом из контента встаем в его конец
		this.Content[Count - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(Count - 1, Depth);
		SearchPos.Found		= true;
		SearchPos.UpdatePos	= true;
		return true;
	}
};
CMathContent.prototype.Get_StartPos = function(ContentPos, Depth)
{
    ContentPos.Update(0, Depth);
    this.Content[0].Get_StartPos(ContentPos, Depth + 1);
};
CMathContent.prototype.Get_EndPos = function(BehindEnd, ContentPos, Depth)
{
    var nLastPos = this.Content.length - 1;

    ContentPos.Update(nLastPos, Depth);
    if(undefined !== this.Content[nLastPos])
        this.Content[nLastPos].Get_EndPos(BehindEnd, ContentPos, Depth + 1);
};
CMathContent.prototype.Draw_HighLights = function(PDSH, bAll)
{
	if (!this.bRoot && this.Parent && true !== this.ParentElement.Is_ContentUse(this))
		return;

    var Bound = this.Get_LineBound(PDSH.Line, PDSH.Range);
    PDSH.X    = Bound.X;

    var CurLine  = PDSH.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PDSH.Range - this.StartRange : PDSH.Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var Y0 = PDSH.Y0,
        Y1 = PDSH.Y1;

    var FirstRunInRootNotShd = this.bRoot && this.Content.length > 0 && this.Content[StartPos].IsShade() == false;

    if(FirstRunInRootNotShd || this.bRoot == false)
    {
        Y0 = Bound.Y;
        Y1 = Bound.Y + Bound.H;
    }

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        PDSH.Y0 = Y0;
        PDSH.Y1 = Y1;

        if(bAll && this.Content[CurPos].Type == para_Math_Run)
            this.Content[CurPos].SelectAll();

        this.Content[CurPos].Draw_HighLights(PDSH, bAll);
    }
};
CMathContent.prototype.Draw_Lines = function(PDSL)
{
    var CurLine  = PDSL.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PDSL.Range - this.StartRange : PDSL.Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var Bound = this.Get_LineBound(PDSL.Line, PDSL.Range);

    var Baseline = Bound.Y + Bound.Asc;

    PDSL.Baseline = Baseline;
    PDSL.X        = Bound.X;

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Draw_Lines(PDSL);
        PDSL.Baseline = Baseline;
    }
};
CMathContent.prototype.RemoveSelection = function()
{
    var StartPos = this.Selection.StartPos;
    var EndPos   = this.Selection.EndPos;

    if (StartPos > EndPos)
    {
        StartPos = this.Selection.EndPos;
        EndPos   = this.Selection.StartPos;
    }

    StartPos = Math.max(0, StartPos);
    EndPos   = Math.min(this.Content.length - 1, EndPos);

    for (var nPos = StartPos; nPos <= EndPos; nPos++)
    {
        this.Content[nPos].RemoveSelection();
    }

    this.Selection.Use   = false;
    this.Selection.StartPos = 0;
    this.Selection.EndPos   = 0;
};
CMathContent.prototype.SelectAll = function(Direction)
{
	this.Selection.Use = true;
	
	if (Direction < 0)
	{
		this.Selection.StartPos = this.Content.length - 1;
		this.Selection.EndPos   = 0;
	}
	else
	{
		this.Selection.StartPos = 0;
		this.Selection.EndPos   = this.Content.length - 1;
	}

    for (var nPos = 0, nCount = this.Content.length; nPos < nCount; nPos++)
    {
        this.Content[nPos].SelectAll(Direction);
    }
};
CMathContent.prototype.Selection_DrawRange = function(_CurLine, _CurRange, SelectionDraw)
{
    var SelectionStartPos = this.Selection.StartPos;
    var SelectionEndPos   = this.Selection.EndPos;

    if(SelectionStartPos > SelectionEndPos)
    {
        SelectionStartPos = this.Selection.EndPos;
        SelectionEndPos   = this.Selection.StartPos;
    }

    var SelectionUse = this.Selection.Use;

    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    if(this.bRoot == false)
    {
        var Bound = this.Get_LineBound(_CurLine, _CurRange);
        SelectionDraw.StartY   = Bound.Y;
        SelectionDraw.H        = Bound.H;
    }

    for(var CurPos = StartPos; CurPos <= EndPos; CurPos++)
    {
        var Item = this.Content[CurPos];

        var bSelectAll = SelectionUse && SelectionStartPos <= CurPos && CurPos <= SelectionEndPos && SelectionStartPos !== SelectionEndPos;

        if(Item.Type == para_Math_Composition && bSelectAll)
        {
            SelectionDraw.FindStart = false;
            SelectionDraw.W += Item.GetWidth(_CurLine, _CurRange);
        }
        else
        {
            Item.Selection_DrawRange( _CurLine, _CurRange, SelectionDraw );
        }
    }
};
CMathContent.prototype.SelectElementByPos = function(nPos)
{
	this.Selection.Use      = true;
	this.Selection.StartPos = nPos;
	this.Selection.EndPos   = nPos;

	this.Content[nPos].SelectAll();

	this.Correct_Selection();

	if (!this.bRoot)
		this.ParentElement.Select_MathContent(this);
	else
		this.ParaMath.bSelectionUse = true;
};
CMathContent.prototype.Select_Element = function(Element, bWhole)
{
    var nPos = -1;
    for(var nCurPos = 0, nCount = this.Content.length; nCurPos < nCount; nCurPos++)
    {
        if (this.Content[nCurPos] === Element)
        {
            nPos = nCurPos;
            break;
        }
    }

    if (-1 !== nPos)
    {
        this.Selection.Use   = true;
        this.Selection.StartPos = nPos;
        this.Selection.EndPos   = nPos;

        if (bWhole)
            this.Correct_Selection();

        if (!this.bRoot)
            this.ParentElement.Select_MathContent(this);
        else
            this.ParaMath.bSelectionUse = true;
    }
};
CMathContent.prototype.SetCurrentElement = function(oElement)
{
	for(var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		if (this.Content[nPos] === oElement)
		{
			this.CurPos = nPos;

			if (!this.bRoot)
				this.ParentElement.SetCurrentMathContent(this);

			break;
		}
	}
};
CMathContent.prototype.Correct_Selection = function()
{
    if (true !== this.Selection.Use)
        return;

    // Здесь мы делаем так, чтобы селект всегда начинался и заканчивался в ране.
    // Предполагается, что контент скорректирован верно до выполнения данной функции.

    var nContentLen = this.Content.length;
    var nStartPos = Math.max(0, Math.min(this.Selection.StartPos, nContentLen - 1));
    var nEndPos   = Math.max(0, Math.min(this.Selection.EndPos,   nContentLen - 1));

    if (nStartPos > nEndPos)
    {
        var nTemp = nStartPos;
        nStartPos = nEndPos;
        nEndPos   = nTemp;
    }

    var oStartElement = this.Content[nStartPos];
    if (para_Math_Run !== oStartElement.Type)
    {
        // Предыдущий элемент должен быть раном
        this.Selection.StartPos = nStartPos - 1;
        this.Content[this.Selection.StartPos].Set_SelectionAtEndPos();
    }

    var oEndElement = this.Content[nEndPos];
    if (para_Math_Run !== oEndElement.Type)
    {
        // Следующий элемент должен быть раном
        this.Selection.EndPos = nEndPos + 1;
        this.Content[this.Selection.EndPos].Set_SelectionAtStartPos();
    }
};
CMathContent.prototype.Create_FontMap = function(Map)
{
    // ArgSize компилируется только тогда, когда выставлены все ссылки на родительские классы
    for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; nIndex++)
        this.Content[nIndex].Create_FontMap(Map);
};
CMathContent.prototype.Get_AllFontNames = function(AllFonts)
{
    for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; nIndex++)
        this.Content[nIndex].Get_AllFontNames(AllFonts);
};
CMathContent.prototype.Selection_CheckParaContentPos = function(ContentPos, Depth, bStart, bEnd)
{
    var CurPos = ContentPos.Get(Depth);

    var bStartPos = this.Selection.StartPos,
        bEndPos   = this.Selection.EndPos;

    if(bStartPos > bEndPos)
    {
        var temp = bStartPos;
        bStartPos = bEndPos;
        bEndPos = temp;
    }

    if(bStartPos < CurPos)
        bStart = false;

    if(CurPos < bEndPos)
        bEnd = false;


    if(bStart === false && bEnd === false)
        return true;
    else if((bStartPos <= CurPos || bStart === false) && (CurPos <= bEndPos || bEnd === false))
        return this.Content[CurPos].Selection_CheckParaContentPos(ContentPos, Depth + 1, bStart, bEnd);

    return false;
};
CMathContent.prototype.Check_NearestPos = function(ParaNearPos, Depth)
{
    var HyperNearPos = new CParagraphElementNearPos();
    HyperNearPos.NearPos = ParaNearPos.NearPos;
    HyperNearPos.Depth   = Depth;

    this.NearPosArray.push(HyperNearPos);
    ParaNearPos.Classes.push(this);

    var CurPos = ParaNearPos.NearPos.ContentPos.Get(Depth);
    this.Content[CurPos].Check_NearestPos(ParaNearPos, Depth + 1);
};
CMathContent.prototype.Recalculate_Range = function(PRS, ParaPr, Depth)
{
    this.bOneLine = PRS.bMath_OneLine;

    // для неинлайн формул:
    // у операторов, находяхщихся на этом уровне (в Run) приоритет выше, чем у внутренних операторов (внутри мат объектов)
    // возможен только принудительный разрыв
    var bOnlyForcedBreak = PRS.bOnlyForcedBreak;
    var bNoOneBreakOperator = PRS.bNoOneBreakOperator;

    var CurLine  = PRS.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PRS.Range - this.StartRange : PRS.Range );

    var ContentLen = this.Content.length;

    var RangeStartPos = this.protected_AddRange(CurLine, CurRange);
    var RangeEndPos   = ContentLen - 1;

    if(this.RecalcInfo.bEqArray)
        this.InfoPoints.SetDefault();

    var ascent = 0, descent = 0;
    this.size.width = 0;
    var bInline = this.ParaMath.Is_Inline();
    var bOperBefore = this.ParaMath.Is_BrkBinBefore(); // true - оператор находится в начале строки, false - оператор находится в конце строки

    // для внутристроковой формулы : начало формулы - начало нового слова
    if(this.bRoot && bInline && true == this.IsStartRange(PRS.Line, PRS.Range) && PRS.Ranges.length == 0)
    {
        PRS.Update_CurPos(0, Depth);
        PRS.Update_CurPos(0, Depth+1); // нулевой элемент всегда Run

        PRS.Set_LineBreakPos(0);

        if(PRS.Word == true)
        {
            PRS.FirstItemOnLine = false;
            PRS.X += PRS.SpaceLen + PRS.WordLen;

			PRS.Word               = false;
			PRS.EmptyLine          = false;
			PRS.EmptyLineWithBreak = false;
			PRS.SpaceLen           = 0;
			PRS.WordLen            = 0;
			PRS.TextOnLine         = true;

            PRS.XRange = PRS.X;
        }
    }

    var bCurInsideOper   = false;

    for(var Pos = RangeStartPos; Pos < ContentLen; Pos++)
    {
        var Item = this.Content[Pos],
            Type = Item.Type;

        PRS.bEndRunToContent = Pos == ContentLen - 1;

        // для однострочных мат объектов обновляем CurLine и CurRange, Run в этом случае не могут разбиваться на несколько строк
        if (this.bOneLine || (0 === Pos && 0 === CurLine && 0 === CurRange ) || Pos !== RangeStartPos)
            Item.Recalculate_Reset( PRS.Range, PRS.Line, PRS );

        PRS.Update_CurPos( Pos, Depth );

        if(this.bOneLine == true) // контент занимает всегда(!) одну строку
        {
            Item.Recalculate_Range_OneLine(PRS, ParaPr, Depth + 1);

            if(this.RecalcInfo.bEqArray && Type == para_Math_Composition)
                this.InfoPoints.ContentPoints.UpdatePoint(this.Content[Pos].size.width);

            this.size.width += Item.size.width;

            if(ascent < Item.size.ascent)
                ascent = Item.size.ascent;

            if(descent < Item.size.height - Item.size.ascent)
                descent = Item.size.height - Item.size.ascent;

            this.size.ascent = ascent;
            this.size.height = ascent + descent;
        }
        else if(bOnlyForcedBreak == true)
        {
            if(Type == para_Math_Run)
            {
                if(true === Item.Is_StartForcedBreakOperator())
                {
                    Item.Recalculate_Range(PRS, ParaPr, Depth + 1);
                }
                else
                {
                    Item.Recalculate_Range_OneLine(PRS, ParaPr, Depth + 1);
                    PRS.WordLen += Item.size.width;
                }
            }
            else
            {
                Item.Recalculate_Range(PRS, ParaPr, Depth + 1);

                if(Item.kind == MATH_BOX && true == Item.IsForcedBreak())
                {
                    this.private_ForceBreakBox(PRS, Item.size.width);
                }
            }

            if(PRS.NewRange == false && PRS.X + PRS.WordLen + PRS.SpaceLen > PRS.XEnd)
            {
                PRS.NewRange = true;
                PRS.MoveToLBP = true;
            }

        }
        else // контент может занимать несколько строк
        {
            // запомним позицию конца Run перед тем как зайдем во внутр мат объект (конечная позиция Run может измениться при пересчете внутр мат объекта)
            var _Depth      = PRS.PosEndRun.Depth;
            var PrevLastPos = PRS.PosEndRun.Get(_Depth-1),
                LastPos     = PRS.PosEndRun.Get(_Depth);

            var PrevWord = PRS.Word,
                MathFirstItem = PRS.MathFirstItem;

            PRS.bInsideOper = false;

            Item.Recalculate_Range(PRS, ParaPr, Depth + 1);

            if(Type == para_Math_Composition)
            {
                // перед мат объектом идет box break_operator и он не является первым элементом в строке
                if(Item.kind == MATH_BOX)
                {
                    PRS.MathFirstItem = MathFirstItem;
                    if(true == Item.IsForcedBreak())
                    {
                        this.private_ForceBreakBox(PRS, Item, _Depth, PrevLastPos, LastPos); // _Depth, PrevLastPos, LastPos запоминаем до пересчета, поэтому передаем эти параметры в функцию
                        PRS.bBreakBox = true;
                    }
                    else if(true == Item.IsOperatorEmulator())
                    {
                        this.private_BoxOperEmulator(PRS, Item, _Depth, PrevLastPos, LastPos); // _Depth, PrevLastPos, LastPos запоминаем до пересчета, поэтому передаем эти параметры в функцию
                        PRS.bBreakBox = true;
                    }
                    else
                    {
                        PRS.WordLen += Item.size.width;
                        PRS.Word = true;

                        if(PRS.X + PRS.SpaceLen + PRS.WordLen > PRS.XEnd)
                        {
                            if (PRS.FirstItemOnLine == false)
                            {
                                PRS.MoveToLBP = true;
                                PRS.NewRange = true;

                                this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
                            }
                            else
                            {
                                PRS.bMathWordLarge = true;
                            }
                        }
                    }

                    PRS.MathFirstItem = false;
                }
                else
                {
                    // Слово не убирается в отрезке. Переносим слово в следующий отрезок
                    // FirstItemOnLine == false - слово оказалось не единственным элементом в промежутке, делаем перенос
                    if(PRS.X + PRS.SpaceLen + PRS.WordLen > PRS.XEnd)
                    {
                        if (PRS.FirstItemOnLine == false)
                        {
                            PRS.MoveToLBP = true;
                            PRS.NewRange = true;

                            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
                        }
                        else
                        {
                            PRS.bMathWordLarge = true;
                        }
                    }


                    // обновляем BreakPos на конец Run, т.к. внутри мат объекта BreakPos  может измениться на  if(true !== Word)
                    // обновляем только в том случае, если Word = false, иначе можем здесь перебить корректный LineBreakPos
                    if(bCurInsideOper == true && PrevWord == false && bOperBefore == false && bNoOneBreakOperator == false && PRS.bInsideOper == false)
                    {
                        // обновим : начало нового слова - конец предыдущего Run

                        PRS.Update_CurPos(PrevLastPos, _Depth-1);
                        PRS.Set_LineBreakPos(LastPos);

                        if(PRS.NewRange == true) // делаем возврат к предыдущему оператору
                            PRS.MoveToLBP = true;
                    }

                    PRS.Word = true;

                }

                if(PRS.NewRange == false && 0 === CurRange && 0 === CurLine)
                {
                    var PrevRecalcInfo = PRS.RunRecalcInfoLast,
                        NumberingAdd;

                    if(null === PrevRecalcInfo)
                        NumberingAdd = true;
                    else
                        NumberingAdd = PrevRecalcInfo.NumberingAdd;

                    if(NumberingAdd)
                    {
                        PRS.X = PRS.Recalculate_Numbering(Item, this, ParaPr, PRS.X);
                        PRS.RunRecalcInfoLast.NumberingAdd = false;
                        PRS.RunRecalcInfoLast.NumberingUse = true;
                        PRS.RunRecalcInfoLast.NumberingItem = PRS.Paragraph.Numbering;
                    }
                }
            }
            else
            {
                if(PRS.MathFirstItem == true && false == Item.IsEmptyRange(PRS.Line, PRS.Range))
                {
                    PRS.MathFirstItem = false;
                }

                var CheckWrapIndent = PRS.bFirstLine == true ? PRS.X - PRS.XRange > PRS.WrapIndent : true;

                if(PRS.bInsideOper == true && CheckWrapIndent == true)
                {
                    PRS.bOnlyForcedBreak = true;
                }
            }

            bCurInsideOper = bCurInsideOper || PRS.bInsideOper;
        }

        if ( true === PRS.NewRange )
        {
            RangeEndPos = Pos;
            break;
        }
    }

    PRS.bInsideOper = PRS.bInsideOper || bCurInsideOper;
    PRS.bOnlyForcedBreak = bOnlyForcedBreak;

    if ( Pos >= ContentLen )
    {
        RangeEndPos = Pos - 1;
    }

    var bSingleBarFraction = false;

    for(var Pos = 0; Pos < ContentLen; Pos++)
    {

        if(this.Content[Pos].kind == MATH_FRACTION && this.Content[Pos].Pr.type == BAR_FRACTION)
        {
            if(bSingleBarFraction)
            {
                bSingleBarFraction = false;
                break;
            }

            bSingleBarFraction = true;
        }
        else if( !(this.Content[Pos].Type == para_Math_Run && true == this.Content[Pos].Is_Empty()) ) // не пустой Run, другой мат объект
        {
            bSingleBarFraction = false;
            break;
        }
    }

    PRS.bSingleBarFraction = bSingleBarFraction;

    this.protected_FillRange(CurLine, CurRange, RangeStartPos, RangeEndPos);
};
CMathContent.prototype.private_ForceBreakBox = function(PRS, Box, _Depth, PrevLastPos, LastPos)
{
    var BoxLen = Box.size.width;

    if(true == PRS.MathFirstItem)
    {
        PRS.WordLen += BoxLen;
    }
    else
    {
        if(true === this.ParaMath.Is_BrkBinBefore())
        {
            PRS.X += PRS.SpaceLen + PRS.WordLen;
            PRS.Update_CurPos(PrevLastPos, _Depth-1);
            PRS.Set_LineBreakPos(LastPos);

            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);

            PRS.MoveToLBP = true;
            PRS.NewRange = true;
        }
        else
        {
            PRS.Word         = false;
            PRS.bForcedBreak = true;
            PRS.WordLen     += BoxLen;
        }
    }
};
CMathContent.prototype.private_BoxOperEmulator = function(PRS, Box, _Depth, PrevLastPos, LastPos)
{
    var BoxLen      = Box.size.width,
        BoxGapRight = Box.GapRight;

    var CheckWrapIndent = PRS.bFirstLine == true ? PRS.X - PRS.XRange > PRS.WrapIndent : true;
    var bOperBefore = this.ParaMath.Is_BrkBinBefore() == true;

    var bOnlyForcedBreakBefore = bOperBefore == true && PRS.MathFirstItem == false,
        bOnlyforcedBreakAfter  = bOperBefore == false;

    if(CheckWrapIndent == true && (bOnlyForcedBreakBefore || bOnlyforcedBreakAfter))
    {
        PRS.bOnlyForcedBreak = true;
    }

    var bOverXEnd;

    if(bOperBefore) // оператор находится в начале строки
    {
        bOverXEnd = PRS.X + PRS.WordLen + PRS.SpaceLen > PRS.XEnd;

        if(true == PRS.MathFirstItem)
        {
            PRS.WordLen += PRS.SpaceLen + PRS.WordLen + BoxLen;
        }
        else if(PRS.FirstItemOnLine == false && bOverXEnd)
        {
            PRS.MoveToLBP = true;
            PRS.NewRange = true;

            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
        }
        else
        {
            PRS.X += PRS.SpaceLen + PRS.WordLen;
            // обновим : начало нового слова - конец предыдущего Run
            PRS.bInsideOper = true;
            PRS.FirstItemOnLine = false;

            PRS.Update_CurPos(PrevLastPos, _Depth-1);
            PRS.Set_LineBreakPos(LastPos);

            PRS.SpaceLen = BoxLen;

            PRS.WordLen = 0;
            PRS.Word = true;
        }
    }
    else
    {
        bOverXEnd = PRS.X + PRS.SpaceLen + PRS.WordLen + BoxLen - BoxGapRight > PRS.XEnd;

        PRS.OperGapRight = BoxGapRight;

        // Слово не убирается в отрезке. Переносим слово в следующий отрезок
        // FirstItemOnLine == false - слово оказалось не единственным элементом в промежутке, делаем перенос
        if(PRS.FirstItemOnLine == false && bOverXEnd)
        {
            PRS.MoveToLBP = true;
            PRS.NewRange = true;

            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
        }
        else
        {
            PRS.bInsideOper = true;
        }

        PRS.X += PRS.SpaceLen + PRS.WordLen + BoxLen;
        PRS.SpaceLen = 0;
        PRS.WordLen = 0;

        PRS.Word = false;
        PRS.FirstItemOnLine = false;
    }
};
CMathContent.prototype.Math_Set_EmptyRange = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var RangeStartPos = this.protected_AddRange(CurLine, CurRange);
    var RangeEndPos   = RangeStartPos;

    this.protected_FillRange(CurLine, CurRange, RangeStartPos, RangeEndPos);

    this.Content[RangeStartPos].Math_Set_EmptyRange(_CurLine, _CurRange);
};
CMathContent.prototype.Recalculate_Reset = function(StartRange, StartLine, PRS)
{
    var bNotUpdate = PRS !== null && PRS!== undefined && PRS.bFastRecalculate == true;
    if(bNotUpdate == false)
    {
        this.StartLine   = StartLine;
        this.StartRange  = StartRange;

        if(this.Content.length > 0)
            this.Content[0].Recalculate_Reset(StartRange, StartLine, PRS);

        this.protected_ClearLines();
    }
};
CMathContent.prototype.IsEmptyRange = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var bEmpty = true;

    if(StartPos == EndPos)
    {
        bEmpty = this.Content[StartPos].IsEmptyRange(_CurLine, _CurRange);
    }
    else
    {
        var Pos = StartPos;

        while(Pos <= EndPos)
        {
            if(false == this.Content[Pos].IsEmptyRange(_CurLine, _CurRange))
            {
                bEmpty = false;
                break;
            }
            Pos++;
        }
    }

    return bEmpty;
};
CMathContent.prototype.Displace_BreakOperator = function(isForward, bBrkBefore, CountOperators)
{
	var Pos = this.CurPos;

	if (this.Content[Pos].Type == para_Math_Run)
	{
		var bApplyBreak = this.Content[Pos].Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
		var NewPos      = bBrkBefore ? Pos + 1 : Pos - 1;

		if (this.Content[NewPos] && bApplyBreak == false && (this.Content[NewPos].Type == para_Math_Run || this.Content[NewPos].kind == MATH_BOX))
		{
			this.Content[NewPos].Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
		}
	}
	else
	{
		this.Content[Pos].Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
	}
};
CMathContent.prototype.Recalculate_Range_Width = function(PRSC, _CurLine, _CurRange)
{
    var RangeW = PRSC.Range.W;
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Recalculate_Range_Width( PRSC, _CurLine, _CurRange );
    }

    this.Bounds.SetWidth(CurLine, CurRange, PRSC.Range.W - RangeW);
};
CMathContent.prototype.recalculateAllSize = function(textMeasurer)
{
	var ascent = 0, descent = 0;
	this.size.width = 0;
	
	if (this.RecalcInfo.bEqArray)
		this.InfoPoints.SetDefault();
	
	for (let pos = 0, count = this.Content.length; pos < count; ++pos)
	{
		let item = this.Content[pos];
		if (para_Math_Run === item.Type)
		{
			item.Math_RecalculateContent();
		}
		else
		{
			item.recalculateAllSize(textMeasurer);
			
			if (this.RecalcInfo.bEqArray)
				this.InfoPoints.ContentPoints.UpdatePoint(item.size.width);
		}
		
		
		this.size.width += item.size.width;
		
		if (ascent < item.size.ascent)
			ascent = item.size.ascent;
		
		if (descent < item.size.height - item.size.ascent)
			descent = item.size.height - item.size.ascent;
		
		this.size.ascent = ascent;
		this.size.height = ascent + descent;
	}
};
CMathContent.prototype.RecalculateMinMaxContentWidth = function(MinMax)
{
	for (let pos = 0, count = this.Content.length; pos < count; ++pos)
	{
		this.Content[pos].RecalculateMinMaxContentWidth(MinMax);
	}
};
CMathContent.prototype.Recalculate_LineMetrics = function(PRS, ParaPr, _CurLine, _CurRange, ContentMetrics)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = (0 === CurLine ? _CurRange - this.StartRange : _CurRange);

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    if(PRS.bFastRecalculate === false)
    {
        this.Bounds.Reset(CurLine, CurRange);
    }

    var NewContentMetrics = new CMathBoundsMeasures();

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        var Item = this.Content[Pos];
        Item.Recalculate_LineMetrics(PRS, ParaPr, _CurLine, _CurRange, NewContentMetrics);
    }

    this.Bounds.UpdateMetrics(CurLine, CurRange, NewContentMetrics);

    ContentMetrics.UpdateMetrics(NewContentMetrics);
};
CMathContent.prototype.Math_UpdateLineMetrics = function(PRS, ParaPr)
{
    this.Content[0].Math_UpdateLineMetrics(PRS, ParaPr); // this.Content[0] - Run
};
CMathContent.prototype.UpdateOperators = function(_CurLine, _CurRange, bEmptyGapLeft, bEmptyGapRight)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        var _bEmptyGapLeft  = bEmptyGapLeft  && Pos == StartPos,
            _bEmptyGapRight = bEmptyGapRight && Pos == EndPos;

        this.Content[Pos].UpdateOperators(_CurLine, _CurRange, _bEmptyGapLeft, _bEmptyGapRight);
    }
};
CMathContent.prototype.Get_Bounds = function()
{
    return this.Bounds.Get_Bounds();
};
CMathContent.prototype.Get_LineBound = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.Get_LineBound(CurLine, CurRange);
};
CMathContent.prototype.GetPos = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetPos(CurLine, CurRange);
};
CMathContent.prototype.GetWidth = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetWidth(CurLine, CurRange);
};
CMathContent.prototype.GetAscent = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetAscent(CurLine, CurRange);
};
CMathContent.prototype.GetDescent = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetDescent(CurLine, CurRange);
};
CMathContent.prototype.Get_StartRangePos = function(_CurLine, _CurRange, SearchPos, Depth, bStartPos)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var CurPos = this.CurPos;
    var Result;

    var bStart = this.bRoot ? CurRange == 0 : bStartPos;

    if(this.Content[CurPos].Type == para_Math_Composition && bStartPos !== true)
    {
        bStart = bStart && CurPos == StartPos;
        Result = this.Content[CurPos].Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bStart); // пытаемся встать в начало внутреннего контента

        if ( true === Result )
        {
            SearchPos.Pos.Update( CurPos, Depth );
        }
        else if(this.bRoot && CurPos !== StartPos)
        {
            if(this.Content[StartPos].Type == para_Math_Composition)
                Result = this.Content[StartPos].Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);
            else
                Result = this.Content[StartPos].Math_Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);

            if ( true === Result )
                SearchPos.Pos.Update(StartPos, Depth);
        }
    }
    else
    {
        if(this.bRoot && CurLine == 0)
        {
            bStart = bStart && StartPos < CurPos;
        }

        if(this.Content[StartPos].Type == para_Math_Composition)
            Result = this.Content[StartPos].Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bStart); // может произойти, если мат объект разбивается на строки
        else
            Result = this.Content[StartPos].Math_Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bStart);

        if ( true === Result )
            SearchPos.Pos.Update(StartPos, Depth );
    }

    return Result;
};
CMathContent.prototype.Get_EndRangePos = function(_CurLine, _CurRange, SearchPos, Depth, bEndPos)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var EndPos = this.protected_GetRangeEndPos(CurLine, CurRange);
    var CurPos = this.CurPos;
    var Result;

    var bLastRange = CurRange == this.protected_GetRangesCount(CurLine) - 1;

    var bEnd = this.bRoot ? bLastRange : bEndPos;

    if(this.Content[CurPos].Type == para_Math_Composition && bEndPos !== true)
    {
        Result = this.Content[CurPos].Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1); // пытаемся встать в конец внутреннего контента

        if ( true === Result )
        {
            SearchPos.Pos.Update( CurPos, Depth );
        }
        else if(this.bRoot && CurPos !== EndPos)
        {
            if(this.Content[EndPos].Type == para_Math_Composition)
                Result = this.Content[EndPos].Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);
            else
                Result = this.Content[EndPos].Math_Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);

            if ( true === Result )
                SearchPos.Pos.Update(EndPos, Depth );
        }
    }
    else
    {
        bEnd = this.bRoot && bLastRange ? true /*иначе после того как встанем после формулы перед знаком параграфа, на следующем End встанем перед формулой*/ : bEnd || CurPos < EndPos;

        if(this.Content[EndPos].Type == para_Math_Composition)
            Result = this.Content[EndPos].Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bEnd); // может произойти, если мат объект разбивается на строки
        else
            Result = this.Content[EndPos].Math_Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bEnd);

        if ( true === Result )
            SearchPos.Pos.Update(EndPos, Depth );

    }

    return Result;
};
CMathContent.prototype.Math_Is_End = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var EndPos = this.protected_GetRangeEndPos(CurLine, CurRange);

    var Len = this.Content.length;

    var result = false;

    if(EndPos == Len - 1 && EndPos >= 0) //TODO check: EndPos >= 0
    {
        result = this.Content[Len - 1].Math_Is_End(_CurLine, _CurRange);
    }

    return result;
};
CMathContent.prototype.Get_AlignBrk = function(_CurLine, bBrkBefore)
{
    var AlnAt = null;

    var CurLine  = _CurLine - this.StartLine;
    var RangesCount = this.protected_GetRangesCount(CurLine - 1);
    var EndPos = this.protected_GetRangeEndPos(CurLine - 1, RangesCount - 1);

    if(CurLine !== 0) // получаем смещение до расчета Recalculate_Range
    {
        var bEndRun  = this.Content[EndPos].Type == para_Math_Run && true == this.Content[EndPos].Math_Is_End(_CurLine - 1, RangesCount - 1),
            bNextBox = EndPos < this.Content.length - 1 && this.Content[EndPos + 1].kind == MATH_BOX;

        var bCheckNextBox = bEndRun == true && bNextBox == true && bBrkBefore == true;

        var bRunEmptyRange = this.Content[EndPos].Type == para_Math_Run && this.Content[EndPos].IsEmptyRange(_CurLine - 1, RangesCount - 1),
            bPrevBox = EndPos > 0 && this.Content[EndPos - 1].kind == MATH_BOX;

        var bCheckPrevNextBox = bRunEmptyRange == true && bPrevBox == true && bBrkBefore == false;

        if(bCheckPrevNextBox)
        {
            AlnAt = this.Content[EndPos - 1].Get_AlignBrk(_CurLine, bBrkBefore);
        }
        else if(bCheckNextBox)
        {
            AlnAt = this.Content[EndPos + 1].Get_AlignBrk(_CurLine, bBrkBefore);
        }
        else
        {
            AlnAt = this.Content[EndPos].Get_AlignBrk(_CurLine, bBrkBefore);
        }
    }

    return AlnAt;
};
CMathContent.prototype.IsStartRange = function(Line, Range)
{
    return Line - this.StartLine == 0 && Range - this.StartRange == 0;
};
CMathContent.prototype.IsStartLine = function(Line)
{
    return Line == this.StartLine;
};
CMathContent.prototype.GetSelectDirection = function()
{
    if (true !== this.Selection.Use)
        return 0;

    if (this.Selection.StartPos < this.Selection.EndPos)
        return 1;
    else if (this.Selection.StartPos > this.Selection.EndPos)
        return -1;

    return this.Content[this.Selection.StartPos].GetSelectDirection();
};
CMathContent.prototype.MoveCursorToStartPos = function()
{
    this.CurPos = 0;
    this.Content[0].MoveCursorToStartPos();
};
CMathContent.prototype.MoveCursorToEndPos = function(SelectFromEnd)
{
    this.CurPos = this.Content.length - 1;
    this.Content[this.CurPos].MoveCursorToEndPos(SelectFromEnd);
};
CMathContent.prototype.Check_Composition = function()
{
    var Pos = this.private_FindCurrentPosInContent();

    return Pos !== null && this.Content[Pos].Type == para_Math_Composition;
};
CMathContent.prototype.Can_ModifyForcedBreak = function(Pr)
{
    var Pos = this.private_GetPosRunForForcedBreak();

    if(Pos !== null && this.bOneLine == false)
    {
        var bBreakOperator      = this.Content[Pos].Check_ForcedBreak();
        var CurrentRun          = this.Content[Pos];
        var bCanCheckNearsRun   = bBreakOperator == false && false == CurrentRun.IsSelectionUse();
        var bPrevItem           = bCanCheckNearsRun && Pos > 0 && true == CurrentRun.Cursor_Is_Start(),
            bNextItem           = bCanCheckNearsRun && Pos < this.Content.length - 1 && true == CurrentRun.Cursor_Is_End();

        var bPrevRun = bPrevItem &&  this.Content[Pos - 1].Type == para_Math_Run,
            bNextRun = bNextItem &&  this.Content[Pos + 1].Type == para_Math_Run;

        if(bBreakOperator)
        {
            this.Content[Pos].Math_Can_ModidyForcedBreak(Pr);
        }
        else if(bPrevRun)
        {
            this.Content[Pos - 1].Math_Can_ModidyForcedBreak(Pr, true, false);
        }
        else if(bNextRun)
        {
            this.Content[Pos + 1].Math_Can_ModidyForcedBreak(Pr, false, true);
        }
    }
};
CMathContent.prototype.private_GetPosRunForForcedBreak = function()
{
    var Pos = null;

    if(true === this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos,
            EndPos   = this.Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }


        var bHaveSelectedItem = false;
        for(var CurPos = StartPos; CurPos <= EndPos; CurPos++)
        {
            var Item = this.Content[CurPos];
            var bSelect     = true !== Item.IsSelectionEmpty(),
                bSelectRun  = bSelect == true && Item.Type == para_Math_Run,
                bSelectComp = bSelect == true && Item.Type == para_Math_Composition;
            var bSelectManyRuns = bSelectRun && bHaveSelectedItem;

            if(bSelectComp || bSelectManyRuns)
            {
                Pos = null;
                break;
            }

            if(bSelectRun)
            {
                bHaveSelectedItem = true;
                Pos = CurPos;
            }
        }
    }
    else
    {
        Pos = this.CurPos;
    }

    return Pos;
};
CMathContent.prototype.private_FindCurrentPosInContent = function()
{
    var Pos = null;

    if(true === this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos,
            EndPos   = this.Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        var bComposition = false;

        for(var CurPos = StartPos; CurPos <= EndPos; CurPos++)
        {
            var Item = this.Content[CurPos];

            if(Item.Type == para_Math_Run && true !== Item.IsSelectionEmpty())
            {
                Pos = bComposition == true ? null : CurPos;
                break;
            }
            else if(Item.Type ==  para_Math_Composition)
            {
                if(bComposition == true)
                {
                    Pos = null;
                    break;
                }
                else
                {
                    Pos = CurPos;
                    bComposition = true;
                }
            }
        }
    }
    else
    {
        this.private_CorrectCurPos();
        Pos = this.CurPos;
    }

    return Pos;
};
CMathContent.prototype.Is_CurrentContent = function()
{
    var Pos = this.private_FindCurrentPosInContent();

    return Pos == null || this.Content[Pos].Type == para_Math_Run;
};
CMathContent.prototype.Set_MenuProps = function(Props)
{
    var Pos = this.private_FindCurrentPosInContent();

    if(true == this.Is_CurrentContent())
    {
        this.Apply_MenuProps(Props, Pos);
    }
    else if(false == this.private_IsMenuPropsForContent(Props.Action) &&  true == this.Content[Pos].Can_ApplyMenuPropsToObject())
    {
        // не нужно проходиться по вложенным элементам
        // 1. уже применили изменения, продожать нет необходимости
        // 2. потому что могут совпать типы текущего элемента и вложенного и тогда изменения применятся к обоим элементам
        if(false === this.Delete_ItemToContentThroughInterface(Props, Pos)) // try to delete
        {
            this.Content[Pos].Apply_MenuProps(Props);
        }
    }
    else
    {
        this.Content[Pos].Set_MenuProps(Props);
    }
};
CMathContent.prototype.Apply_MenuProps = function(Props, Pos)
{
	var ArgSize, NewArgSize;

	if (Props.Action & c_oMathMenuAction.IncreaseArgumentSize)
	{
		if (true === this.Parent.Can_ModifyArgSize() && true == this.Compiled_ArgSz.Can_Increase() && true == this.ArgSize.Can_SimpleIncrease())
		{
			ArgSize    = this.ArgSize.GetValue();
			NewArgSize = this.ArgSize.Increase();

			History.Add(new CChangesMathContentArgSize(this, ArgSize, NewArgSize));
			this.Recalc_RunsCompiledPr();
		}
	}
	else if (Props.Action & c_oMathMenuAction.DecreaseArgumentSize)
	{
		if (true === this.Parent.Can_ModifyArgSize() && true == this.Compiled_ArgSz.Can_Decrease() && true == this.ArgSize.Can_Decrease())
		{
			ArgSize    = this.ArgSize.GetValue();
			NewArgSize = this.ArgSize.Decrease();

			History.Add(new CChangesMathContentArgSize(this, ArgSize, NewArgSize));
			this.Recalc_RunsCompiledPr();
		}
	}

	var Run;

	if (Pos !== null && Props.Action & c_oMathMenuAction.InsertForcedBreak)
	{
		Run = this.private_Get_RunForForcedBreak(Pos);
		Run.Set_MathForcedBreak(true);

	}
	else if (Pos !== null && Props.Action & c_oMathMenuAction.DeleteForcedBreak)
	{
		Run = this.private_Get_RunForForcedBreak(Pos);
		Run.Set_MathForcedBreak(false);
	}
};
CMathContent.prototype.private_Get_RunForForcedBreak = function(Pos)
{
    var CurrentRun          = this.Content[Pos];
    var bCurrentForcedBreak = this.Content[Pos].Type == para_Math_Run && true == CurrentRun.Check_ForcedBreak(),
        bPrevForcedBreak    = Pos > 0 && true == CurrentRun.Cursor_Is_Start(),
        bNextForcedBreak    = Pos < this.Content.length && true == CurrentRun.Cursor_Is_End();

    var Run = null;

    if(bCurrentForcedBreak)
    {
        Run = this.Content[Pos];

        var NewRun = Run.Math_SplitRunForcedBreak();

        if(NewRun !== null)
        {
            this.Internal_Content_Add(Pos+1, NewRun, true);
            Run = NewRun;
        }
    }
    else if(bPrevForcedBreak)
    {
        Run = this.Content[Pos - 1];
    }
    else if(bNextForcedBreak)
    {
        Run = this.Content[Pos + 1];
    }

    return Run;
};
CMathContent.prototype.Delete_ItemToContentThroughInterface = function(Props, Pos)
{
    var bDelete = false;
    var Item = this.Content[Pos];

    if(Item.kind == MATH_DEGREESubSup && Item.Pr.type == DEGREE_SubSup && Props.Type == Asc.c_oAscMathInterfaceType.Script)
    {
        if(Props.ScriptType == Asc.c_oAscMathInterfaceScript.Sup)
        {
            this.private_AddModifiedDegree(Pos, DEGREE_SUPERSCRIPT);
            bDelete = true;
        }
        else if(Props.ScriptType == Asc.c_oAscMathInterfaceScript.Sub)
        {
            this.private_AddModifiedDegree(Pos, DEGREE_SUBSCRIPT);
            bDelete = true;
        }
    }

    var RemoveChar      = Props.Action & c_oMathMenuAction.RemoveAccentCharacter && Item.kind == MATH_ACCENT,
        RemoveBar       = Props.Action & c_oMathMenuAction.RemoveBar && Item.kind == MATH_BAR,
        RemoveScript    = Props.Type == Asc.c_oAscMathInterfaceType.Script && Props.ScriptType == Asc.c_oAscMathInterfaceScript.None && (Item.kind == MATH_DEGREESubSup || Item.kind == MATH_DEGREE),
        RemoveLimit     = Props.Type == Asc.c_oAscMathInterfaceType.Limit && Props.Pos == Asc.c_oAscMathInterfaceLimitPos.None && Item.kind === MATH_LIMIT,
        RemoveMatrix    = Props.Type == Asc.c_oAscMathInterfaceType.Matrix && this.Content[Pos].Is_DeletedItem(Props.Action),
        RemoveEqArray   = Props.Type == Asc.c_oAscMathInterfaceType.EqArray && this.Content[Pos].Is_DeletedItem(Props.Action),
        RemoveDelimiter = Props.Action & c_oMathMenuAction.RemoveDelimiter && Item.kind == MATH_DELIMITER,
        RemoveGroupChar = Props.Type == Asc.c_oAscMathInterfaceType.GroupChar && Props.Pos == Asc.c_oAscMathInterfaceGroupCharPos.None && Item.kind == MATH_GROUP_CHARACTER,
        RemoveRadical   = Props.Action & c_oMathMenuAction.RemoveRadical && Item.kind == MATH_RADICAL;


    if(RemoveChar || RemoveBar || RemoveScript || RemoveLimit || RemoveMatrix || RemoveEqArray || RemoveDelimiter || RemoveGroupChar || RemoveRadical)
    {
        var Items = this.Content[Pos].Get_DeletedItemsThroughInterface();

        if(Items == null) // такого не должно произойти
            return;

        this.Remove_FromContent(Pos, 1);

        this.ConcatToContent(Pos, Items);

        this.Correct_Content();

        bDelete = true;
    }

    return bDelete;
};
CMathContent.prototype.private_AddModifiedDegree = function(Pos, Type)
{
    var DegreeSubSup = this.Content[Pos];
    var Base     = DegreeSubSup.getBase();
    var Iterator = Type == DEGREE_SUBSCRIPT ? DegreeSubSup.getLowerIterator() : DegreeSubSup.getUpperIterator();

    var Degree = new CDegree({type: Type}, false);

    this.Remove_FromContent(Pos, 1);
    this.Add_ToContent(Pos, Degree);

    var oBase     = Degree.getBase(),
        oIterator = Degree.getIterator();

    oBase.Concat_ToEnd(Base.Content);
    oIterator.Concat_ToEnd(Iterator.Content);
};
CMathContent.prototype.Get_MenuProps = function()
{
    var Pr = new CMathMenuBase();

    var Pos = this.private_FindCurrentPosInContent();

    if(Pos !== null && this.Content[Pos].Type == para_Math_Composition)
    {
        Pr = this.Content[Pos].Get_MenuProps();
    }
    else
    {
        this.Can_ModifyForcedBreak(Pr);
    }

    return Pr;
};
CMathContent.prototype.private_IsMenuPropsForContent = function(Action)
{
    // данные изменения могут прийти для любого типа изменений
    var bInsertForcedBreak = Action & c_oMathMenuAction.InsertForcedBreak,
        bDeleteForcedBreak = Action & c_oMathMenuAction.DeleteForcedBreak,
        bIncreaseArgSize   = Action & c_oMathMenuAction.IncreaseArgumentSize,
        bDecreaseArgSize   = Action & c_oMathMenuAction.DecreaseArgumentSize;

    return bDecreaseArgSize || bIncreaseArgSize || bInsertForcedBreak || bDeleteForcedBreak;
};
CMathContent.prototype.MergeParaRuns = function ()
{
	if (this.Content.length > 0) {

        for(let i = 0; i < this.Content.length; i++) {

			if (this.Content[i].Type === 49) {
				let CurrentRun = this.Content[i];
				while (this.Content[i+1] && this.Content[i+1].Type === 49) {
					let oFutureRun = this.Content[i+1];
                    let oContent = oFutureRun.Content.slice(0, oFutureRun.Content.length);
					CurrentRun.ConcatToContent(oContent);
					this.Remove_Content(i + 1, 1);
				}
			}
		}
	}
};
CMathContent.prototype.SplitSelectedContent = function()
{
    if (this.Content.length < 1)
        return;

    if (this.Selection.StartPos > this.Selection.EndPos) {
        let intTemp = this.Selection.StartPos;
        this.Selection.StartPos = this.Selection.EndPos;
        this.Selection.EndPos = intTemp;
    }

    if (this.Selection.StartPos !== this.Selection.EndPos) {
        this.SplitSelectedContentPos(this.Selection.StartPos);
        this.SplitSelectedContentPos(this.Selection.EndPos);
    } else {
        this.SplitSelectedContentPos(this.Selection.StartPos);
    }
};
CMathContent.prototype.SplitSelectedContentPos = function(pos)
{
    let oContent = this.Content[pos];
    let isStart = this.Selection.StartPos === pos;

    if (oContent) {

        if (oContent.Type === 49) {

            if (oContent.Selection.StartPos > oContent.Selection.EndPos) {
                var intTemp = oContent.Selection.StartPos;
                oContent.Selection.StartPos = oContent.Selection.EndPos;
                oContent.Selection.EndPos = intTemp;
            }

            if (oContent.Selection.StartPos !== 0 && oContent.Selection.StartPos <= oContent.Content.length - 1)
            {
                let oPrevContent = oContent.Split_Run(oContent.Selection.StartPos);
                this.Add_ToContent(this.Selection.StartPos + 1, oPrevContent);

                this.Selection.StartPos += 1;
                this.Selection.EndPos += 1;

                if (oPrevContent.Selection.EndPos !== oPrevContent.Content.length) {
                    let oNextContent = oPrevContent.Split_Run(oPrevContent.Selection.EndPos);
                    this.Add_ToContent(this.Selection.EndPos + 1, oNextContent);
                }
            }
            else if (oContent.Selection.StartPos === oContent.Selection.EndPos && isStart)
            {
                this.Selection.StartPos += 1;
            }

            if (oContent.Selection.EndPos !== oContent.Content.length && oContent.Selection.EndPos <= oContent.Content.length - 1)
            {
                let oNextContent = oContent.Split_Run(oContent.Selection.EndPos);
                this.Add_ToContent(this.Selection.EndPos + 1, oNextContent);
            }
            else if (oContent.Selection.StartPos === oContent.Selection.EndPos && !isStart)
            {
                this.Selection.EndPos -= 1;
            }
        }
    }
};
CMathContent.prototype.ConvertContentView = function(intStart, intEnd, nInputType, isToLinear)
{
    if (this.Content.length === 0)
        return;

    if (intStart >= 0 && intEnd <= this.Content.length)
    {
        let strContent = "";
        let intCount = (intEnd - intStart) + 1;

        for (let i = intStart, j = 0; i <= intEnd; i++)
        {
            let oElement = this.Content[i];

            if (undefined !== oElement)
            {
                strContent += oElement.GetTextOfElement(nInputType);
            }
        }

        if ((isToLinear || undefined === nInputType || null === nInputType) && strContent !== "")
        {
            this.Remove_FromContent(intStart, intCount);
            this.Add_TextOnPos(intStart, strContent);

            this.Content[intStart].SelectAll();

            this.Selection.Use      = true;
            this.Selection.StartPos = intStart;
            this.Selection.EndPos   = intStart;
        }
        else
        {
            let oTempContent = new CMathContent();

            if (nInputType === Asc.c_oAscMathInputType.Unicode)
            {
                AscMath.CUnicodeConverter(strContent, oTempContent);
            }
            else if (nInputType === Asc.c_oAscMathInputType.LaTeX)
            {
                AscMath.ConvertLaTeXToTokensList(strContent, oTempContent);
            }

            this.Remove_FromContent(intStart, intCount);
            this.RemoveSelection();

            for (let i = 0; i < oTempContent.Content.length; i++)
            {
                this.Add_ToContent(intStart + i, oTempContent.Content[i], false);
                this.Content[intStart + i].SelectAll();

                if (i === 0)
                {
                    this.Selection.Use      = true;
                    this.Selection.StartPos = intStart +  i;
                    this.Selection.EndPos   = intStart +  i;
                }
                else
                {
                    this.Selection.EndPos   = intStart +  i;
                }
            }

            this.Correct_Content(true)
            this.Correct_Selection();
        }
    }
};
CMathContent.prototype.SplitContentByContentPos = function()
{
    let oCurrentObj = this.Content[this.CurPos];
    let nCursorPos = oCurrentObj.State.ContentPos;
    let arrContent = [];

    if (nCursorPos < oCurrentObj.Content.length)
    {
        if (oCurrentObj.Split_Run)
        {
            let oNewRun = oCurrentObj.Split_Run(nCursorPos);
            arrContent.push(oNewRun);
        }
        else
        {
            // контент в котором мы находимся не является ParaRun
            // значит делить не нужно т.к мы в обертке - выходим и отменяем автокоррекцию
            return false;
        }
    }

    for (let i = this.CurPos + 1; i < this.Content.length; i++)
    {
        arrContent.push(this.Content[i].Copy());
    }

    this.Remove_FromContent(this.CurPos + 1, this.Content.length - this.CurPos - 1);

    return arrContent;
};
CMathContent.prototype.Process_AutoCorrect = function (oElement)
{
    let isConvert = false;
    var oLogicDocument = this.GetLogicDocument();
    var nInputType = oLogicDocument
        ? oLogicDocument. Api.getMathInputType()
        : Asc.c_oAscMathInputType.Unicode;

    // LaTeX autocorrection disabled
    if (nInputType === 1 || oElement.value === 39 || oElement.value === 34)
        return;

    // split content by cursor position
    const arrNextContent = this.SplitContentByContentPos();
	let lastElement = this.GetLastTextElement();

    if (arrNextContent === false)
        return;

    if (this.CorrectSpecialWordOnCursor(nInputType))
    {
        if (arrNextContent)
            this.ConcatToContent(this.Content.length, arrNextContent);

        return;
    }

    // convert content of bracket block, near cursor for Unicode (1/2) -> ( CFraction )
    if (nInputType === 0)
        this.ConvertContentInLastBracketBlock(nInputType);

    // convert word near cursor (\int, \sqrt, \alpha...)
    if (oElement.value === 32 || this.IsLastElement(AscMath.MathLiterals.operators) || lastElement === '(' || lastElement === ")")
    {
        if (oElement.value === 32)
        {
            if (this.CorrectWordOnCursor(nInputType === 1))
            {
                if (arrNextContent)
                    this.ConcatToContent(this.Content.length, arrNextContent);

                return;
            }
        }
        else
        {
            if (this.CorrectWordOnCursor(nInputType === 1, true))
            {
                if (arrNextContent)
                    this.ConcatToContent(this.Content.length, arrNextContent);

                return;
            }
        }
    }

    if (this.IsLastElement(AscMath.MathLiterals.operators))
    {
        let strPreLast = this.GetPreLastTextElement();
        if (strPreLast === "_" || strPreLast === "^")
        {
            return
        }
    }

    // check is needed start autocorrection
    if (!this.IsStartAutoCorrection(nInputType, oElement.value))
    {
        this.AddContentForAutoCorrection(arrNextContent, true);
        return;
    }

    //const oSlashesContent = this.GetSlashesInfo();

    // Unicode
    if (nInputType === 0)
    {
        // proceed bracket block () -> CDelimiter
        let Bracket = this.CheckAutoCorrectionBrackets(nInputType, true);

        // proceed rules (1/2, 1_2 ...)
        isConvert = Bracket.intCounter >= 0 ? this.CheckAutoCorrectionRules(nInputType) : false;

        // else - convert content until first operator
        if (isConvert === false && Bracket.intCounter >= 0 && Bracket.isConvert === false)
            this.CheckWhileOperatorContent(Bracket.OperatorsPos, nInputType, true);
    }
    else // LaTex
    {
        let Bracket = this.CheckAutoCorrectionBrackets(nInputType);
        if (Bracket.intCounter === 0)
            this.CheckWhileOperatorContent(Bracket.OperatorsPos, nInputType, true);
    }

    this.MoveCursorToEndPos();

    if (arrNextContent.length > 0) {
        this.AddContentForAutoCorrection(arrNextContent, true);
    }
};
CMathContent.prototype.GetLastContent = function ()
{
    let oContent = this;
    while (oContent && oContent.Content && oContent.Content.length > 0)
    {
        oContent = oContent.Content[oContent.Content.length - 1];
    }

    return oContent;
}
CMathContent.prototype.GetPreLastContent = function()
{
    let oContent = this;

    while (oContent && oContent.Content && oContent.Content.length > 0)
    {
        if (oContent.Content[oContent.Content.length - 1] &&
            oContent.Content[oContent.Content.length - 1].Content &&
            oContent.Content[oContent.Content.length - 1].Content.length >= 2)
        {
            oContent = oContent.Content[oContent.Content.length - 1];
        }
        else
        {
            oContent = oContent.Content[oContent.Content.length - 2];
        }
    }

    return oContent;
}
CMathContent.prototype.GetLastTextElement = function ()
{
    let oContent = this.GetLastContent();
    if (oContent)
        return String.fromCharCode(oContent.value);
}
CMathContent.prototype.GetPreLastTextElement = function ()
{
    let oContent = this.GetPreLastContent();
    if (oContent)
        return String.fromCharCode(oContent.value);
}
CMathContent.prototype.CheckWhileOperatorContent = function(arrOperatorsList, nInputType, isInEnd)
{
    if (arrOperatorsList === undefined)
        return;

    let arrContentTypes = [];
    let isSpace = false;

    if (this.GetLastTextElement() === " ")
    {
        isSpace = this.DeleteEndSpace();
        for (let i = 0; i < this.Content.length; i++) {
            arrContentTypes.push(this.Content[i].Type);
        }
    }

    if(arrOperatorsList.length === 0)
    {
        this.CutConvertAndPaste([0, 0], nInputType);
    }
    else {
        let Position = arrOperatorsList[arrOperatorsList.length - 1];
        this.CutConvertAndPaste([Position[2], Position[0]], nInputType);
    }

    if (isSpace)
    {
        let prevLen = arrContentTypes.length - 1;
        let isEqual = true;

        for (let i = this.Content.length - 1; i >= 0 && prevLen >= this.Content.length - 1; i--, prevLen--)
        {
            let currentContent = this.Content[i].Type;
            let prevContent = arrContentTypes[prevLen];

            if (prevContent === 49 && currentContent !== 49)
            {
                isEqual = false;
            }

            if (currentContent !== 49)
                break;
        }

        if (isEqual)
           this.Add_TextOnPos(this.Content.length,' ');
    }
};
CMathContent.prototype.DeleteContentForAutoCorrection = function(arrDeleteData)
{
    if (arrDeleteData < 1)
        return;

    var intCounterForDel = 0;

    for (var i = 0; i < arrDeleteData.length; i++)
    {
        if (this.Content.length > 0)
        {
            var intIndex = this.CurPos;
            var oContent = this.Content[intIndex];

            var intLengthOfContent = oContent.Content.length;
            var intDeleteCount = arrDeleteData[intCounterForDel];

            if (intDeleteCount !== undefined && intDeleteCount !== 0)
            {
                if (intLengthOfContent <= intDeleteCount)
                {
                    let intTempCurPos = this.CurPos;
                    this.Remove_FromContent(intIndex, 1);

                    if (intTempCurPos === this.CurPos)
                        this.CurPos--;
                    intCounterForDel++;
                }
                else
                {
                    oContent.Remove_FromContent(intLengthOfContent - intDeleteCount, intDeleteCount);
                    intCounterForDel++;
                }
            }
        }
    }

    if (this.Content.length === 1 && this.Content[0] && this.Content[0].Type === 49 && this.Content[0].Content.length === 0)
    {
        this.Remove_FromContent(0, 1);
    }
};
CMathContent.prototype.AddContentForAutoCorrection = function(arrNewElements, isCurPosChange)
{
    if (arrNewElements.length < 1)
        return;

	this.ConcatToContent( isCurPosChange ? this.CurPos + 1 : this.CurPos, arrNewElements);
};
CMathContent.prototype.CorrectWordOnCursor = function(IsLaTeX, pos)
{
    return AscMath.CorrectWordOnCursor(this, IsLaTeX, pos);
};
CMathContent.prototype.CorrectAllMathWords = function(IsLaTeX)
{
    return AscMath.CorrectAllWords(this, IsLaTeX);
};
CMathContent.prototype.ConvertAllSpecialWords = function (IsLaTeX)
{
    return AscMath.CorrectAllSpecialWords(this, IsLaTeX)
}
CMathContent.prototype.CorrectSpecialWordOnCursor = function (IsLaTeX)
{
    return AscMath.CorrectSpecialWordOnCursor(this, IsLaTeX)
}
CMathContent.prototype.GetBracketOperatorInfo = function (isLaTeX)
{
    const arrContent = {};

    for (let i = 0; i < this.Content.length; i++)
    {
        if (this.Content[i].Type === 49 && this.Content[i].Content.length > 0)
        {
            let cont = this.Content[i].MathAutocorrection_GetBracketsOperatorsInfo(isLaTeX);

            if (cont.length > 0) {
                arrContent[i] = cont;
            }
        }
    }

    return arrContent;
};
CMathContent.prototype.GetOperatorInfo = function ()
{
    const arrContent = {};

    for (let i = 0; i < this.Content.length; i++)
    {
        if (this.Content[i].Type === 49 && this.Content[i].Content.length > 0)
        {
            arrContent[i] = this.Content[i].MathAutocorrection_GetOperatorInfo();
        }
    }

    return arrContent;
};
CMathContent.prototype.GetSlashesInfo = function ()
{
    const arrContent = {};

    for (let i = 0; i < this.Content.length; i++)
    {
        if (this.Content[i].Type === 49 && this.Content[i].Content.length > 0)
        {
            let content = this.Content[i].MathAutocorrection_GetSlashesInfo();
            if (content.length > 0)
            {
                arrContent[i] = content;
            }
        }
    }

    return arrContent;
};
CMathContent.prototype.IsLastElement = function (type)
{
    if (this.Content.length === 0)
        return

    let oLastElement = this.Content[this.Content.length - 1];
    if (oLastElement.Type === 49)
        return oLastElement.MathAutocorrection_IsLastElement(type);

    return false;
};
//авто-конвертации контента ВНУТРИ скобок, не самих скобок
CMathContent.prototype.ConvertContentInLastBracketBlock = function(nInputType)
{
    if (this.IsLastTextElementRBracket())
    {
        const oBracketsContent = this.GetBracketOperatorInfo(nInputType === 1);
        const Brackets = new ProceedBrackets(oBracketsContent);
		let Result = Brackets.GetPosition();

        if (Brackets.intCounter === 0 && Brackets.BracketsPair.length > 0)
        {
            if (Result.length === 2)
            {
                Result[1]++;
                this.CutConvertAndPaste(Result, nInputType);
                Brackets.isConvert = true;
            }
        }
    }
};

function ProceedBrackets(arrDataOfBrackets)
{
    this.Buffer = [];
    this.BracketsPair = [];
    this.OperatorsPos = [];
    this.isContinue = true;
    this.Position = [];
    this.isStopConvertation = false;
    this.intCounter = 0;
    this.isConvert = false;

    let one = Object.keys(arrDataOfBrackets);

    for (let i = one.length - 1; i >= 0 && this.isContinue; i--)
    {
        const BracketsOfCurrentContent = arrDataOfBrackets[one[i]];

        if (!BracketsOfCurrentContent && BracketsOfCurrentContent.length === 0)
            continue;

        this.Position[0] = parseInt(one[i]);

        for (let j = BracketsOfCurrentContent.length - 1; j >= 0 && this.isContinue; j--)
        {
            const intCurrentData = BracketsOfCurrentContent[j];

            if (this.Buffer.length === 0)
                this.PushData(intCurrentData);
            else
                this.AnalyseData(intCurrentData);
        }
    }
}
ProceedBrackets.prototype.GetLastBracketStartPos = function ()
{
	let arrBrack = this.BracketsPair[this.BracketsPair.length - 1][0][0];
	return arrBrack;
}
ProceedBrackets.prototype.PushData = function (intContent)
{
    intContent[2] = this.Position[0]

    if (intContent[1] === 2)
    {
        this.OperatorsPos.push(intContent);
        return
    }
    else if (intContent[1] === -1)
        this.intCounter++;
    else if (intContent[1] === 1)
        this.intCounter--;
    else if (intContent[1] === 0)
    {
        if (this.intCounter > 0)
            this.intCounter++;
        else
            this.intCounter--;
    }

    this.Buffer.push(intContent);
};
ProceedBrackets.prototype.AnalyseData = function (arrContent)
{
    const arrLastElement = this.GetLastData();

    arrContent[2] = this.Position[0];
    if (arrContent[1] === 0)
    {
        if (arrLastElement[1] === 0)
        {
            this.Buffer.shift();
            this.BracketsPair.push([arrContent, arrLastElement]);
            this.intCounter++;
            this.Position[1] = arrContent[0] - 1;
        }
        else if (arrLastElement[1] === 1)
        {
            this.Buffer.shift();
            this.BracketsPair.push([arrContent, arrLastElement]);
            this.intCounter++;
            this.Position[1] = arrContent[0] - 1;
        }
        else
            return;
    }
    else if (arrContent[1] === 1)
    {
        if (arrLastElement[1] === 1)
            this.PushData(arrContent);
        else if (arrLastElement[1] === 2)
            this.OperatorsPos.push(arrLastElement);
        else if (arrLastElement[1] === -1) {
            this.Buffer.shift();
            this.BracketsPair.push([arrContent, arrLastElement]);
            this.intCounter--;
            this.Position[1] = arrContent[0] - 1;
        }
        else
            return;
    }
    else if (arrContent[1] === -1)
    {
        if (arrLastElement[1] === 0)
        {
            this.Buffer.shift();
            this.BracketsPair.push([arrContent, arrLastElement]);
            this.intCounter++;
            this.Position[1] = arrContent[0] - 1;
        }
        else if (arrLastElement[1] === 1)
        {
            this.Buffer.shift();
            this.BracketsPair.push([arrContent, arrLastElement]);
            this.intCounter++;
            this.Position[1] = arrContent[0] - 1;
        }
        else if (arrLastElement[1] === -1)
            return;
        else if (arrLastElement[1] === 2) {
            this.OperatorsPos.push(arrLastElement);
            this.isStopConvertation = true;
        }
        else {
            this.isContinue = false;
            return
        }
    }
};
ProceedBrackets.prototype.GetLastData = function ()
{
    let arrContent = [];
    let intCounter = this.Buffer.length - 1;

    while (arrContent !== 2 && intCounter >= 0)
    {
        arrContent = this.Buffer[intCounter];
        intCounter--;
    }

    if (arrContent)
        return arrContent;

    return false;
};
ProceedBrackets.prototype.IsStop = function()
{
    return this.isStopConvertation === true || this.intCounter !== 0
};
ProceedBrackets.prototype.IsHasPoints = function ()
{
    return this.Position.length === 2;
};
ProceedBrackets.prototype.GetPosition = function ()
{
	let Pair = [];
	let oLastBracket = this.BracketsPair[0];
	if (oLastBracket)
	{
		let oClosePos = oLastBracket[0];
		Pair[0] = oClosePos[2];
		Pair[1] = oClosePos[0];
	}

	return Pair;
};
ProceedBrackets.prototype.IsBufferIncludeOneOpenBracket = function ()
{
    return this.Buffer.length === 1 && this.Buffer[0][1] === -1
};
ProceedBrackets.prototype.FindPairPos = function (intPos)
{
    for (let i = 0; i < this.BracketsPair.length; i++)
    {
        let arrCurrentPair = this.BracketsPair[i];

        if (arrCurrentPair[1] === intPos)
            return arrCurrentPair[0];
    }
    return false;
};

const ParaRunIterator = function(ParaRun)
{
    this.Content = ParaRun.Content;
    this.Cursor = ParaRun.Content.length - 1;
}
ParaRunIterator.prototype.GetNext = function()
{
    if (!this.IsHasContent())
        return false;

    const oContent = this.Content[this.Cursor];
    this.Cursor--;

    return oContent.value;
};
ParaRunIterator.prototype.IsHasContent = function()
{
    return this.Cursor !== -1;
};
ParaRunIterator.prototype.CheckCursor = function()
{
    if (this.Cursor === -1) {
        this.Cursor = 0;
    }
};

const ContentIterator = function(Content, oBracketContent)
{
    this.Content = Content.Content;
    this.cursor = this.Content.length - 1;
    this.RulePosition = [];
    this.BracketPosition = [];
    this.CurrentElement;
    this.intBracketCounter = 0;
    this.BracketContent = oBracketContent;
    this.isBracket = false;
}
ContentIterator.prototype.IsHasContent = function ()
{
    return this.cursor >= 0;
};
ContentIterator.prototype.GetNext = function (nextRule)
{
    if (!this.CurrentElement)
        this.CreateCurrentElement(nextRule);

    return this.GetNextFromCurrentElement(nextRule);
};
ContentIterator.prototype.CreateCurrentElement = function(nextRule)
{
    const oCurrent = this.Content[this.cursor];

    if (oCurrent instanceof ParaRun) {
        if (oCurrent.Content.length === 0) {
            this.cursor--;
            this.CreateCurrentElement(nextRule);
        }
        else
            this.CurrentElement = new ParaRunIterator(oCurrent);
    }
    else
        this.CurrentElement = oCurrent;
};
ContentIterator.prototype.GetNextFromCurrentElement = function (nextRule)
{
    if (this.CurrentElement instanceof ParaRunIterator)
    {
        if (!this.CurrentElement.IsHasContent() && this.cursor > 0)
        {
            this.cursor--;
            this.CreateCurrentElement();
            return this.GetNextFromCurrentElement(nextRule);
        }
        else if (!this.CurrentElement.IsHasContent() && this.cursor <= 0)
        {
            return false;
        }

        let intCount = 0;
        let intCode;

        while (this.CurrentElement.IsHasContent())
        {
            let prevCode = intCode;
            let intCurrent = intCode = this.CurrentElement.GetNext();
            let strCurrent = String.fromCharCode(intCode);

            if (this.CurrentRule !== "\"" && strCurrent === "\"")
            {
                return this.CheckTextLiteral(strCurrent);
            }
            if (AscMath.MathLiterals.rBrackets.IsIncludes(strCurrent)) {
                return this.CheckBracket(strCurrent);
            }
            else if (AscMath.MathLiterals.lBrackets.IsIncludes(strCurrent))
            {
                this.CurrentElement.Cursor++;
                return prevCode;
            }

            else if (strCurrent === "▒" && this.CurrentRule !== "▒") {
                this.CurrentElement.Cursor++;
                return prevCode;
            }

            if (this.CurrentRule === strCurrent)
                return intCurrent;

            else if (nextRule !== true && nextRule && intCurrent === nextRule.charCodeAt(0))
            {
                this.CurrentElement.Cursor++;
                return prevCode;
            }

            if (this.CurrentRule !== true)
                return false;

            intCount++;
        }

        return intCode;
    }
    else
    {
        if (nextRule !== undefined) {
            this.cursor--;
            this.CurrentElement = undefined;
        }

        return true;
    }
};

ContentIterator.prototype.CheckTextLiteral = function(literal)
{
    while (this.CurrentElement)
    {
		if (this.CurrentElement.Cursor < 0)
			return

        if (!this.CurrentElement.IsHasContent() && this.cursor > 0)
        {
            this.cursor--;
            this.CreateCurrentElement();
        }

        if (!(this.CurrentElement instanceof ParaRunIterator) || this.CurrentElement.Cursor < 0)
        {
            this.cursor--;
            this.CreateCurrentElement();
        }
        else if (this.CurrentElement.IsHasContent())
        {
            let intCurrent = this.CurrentElement.GetNext();
            let strCurrent = String.fromCharCode(intCurrent);

            if (strCurrent === literal)
                return true;
        }
    }
}
ContentIterator.prototype.CheckBracket = function(strCurrent)
{
    this.intBracketCounter--;

    while (this.CurrentElement && this.intBracketCounter !== 0)
    {
        if (!this.CurrentElement.IsHasContent() && this.cursor > 0)
        {
            this.cursor--;
            this.CreateCurrentElement();
        }

        if (!(this.CurrentElement instanceof ParaRunIterator))
        {
            this.cursor--;
            this.CreateCurrentElement();
        }

        while (this.CurrentElement.IsHasContent())
        {
            let intCurrent = this.CurrentElement.GetNext();
            let strCurrent = String.fromCharCode(intCurrent);

            if (AscMath.MathLiterals.rBrackets.IsIncludes(strCurrent))
            {
                this.CheckBracket(strCurrent);
            }
            else if (AscMath.MathLiterals.lrBrackets.IsIncludes(strCurrent))
            {
                if (this.intBracketCounter === 0)
                {
                    this.CheckBracket(strCurrent);
                }
                else if (this.intBracketCounter < 0)
                {
                    this.intBracketCounter++;
                }
                else if (this.intBracketCounter > 0)
                {
                    this.intBracketCounter--;
                }
            }
            else if (AscMath.MathLiterals.lBrackets.IsIncludes(strCurrent) && this.intBracketCounter < 0)
            {
                this.intBracketCounter++;
                return true
            }
            else if (AscMath.MathLiterals.lBrackets.IsIncludes(strCurrent))
            {
                this.CurrentElement.CheckCursor();
                this.CurrentElement.Cursor++;
                return
            }
        }
    }
}
ContentIterator.prototype.ResetParaRunCursor = function()
{
    this.cursor = this.Content.length - 1
    this.CurrentElement = undefined;
    this.intBracketCounter = 0;
};
ContentIterator.prototype.CheckRules = function ()
{
	// TODO it is better to search for sequences by what is, and not vice versa; incredibly inefficient!!!
	const rules = [
		//true обозначает обычный текст или блоки контента (CFraction, CLimit, CDegree...);
		// ["_"],
		// ["^"],
		[true, "^", true, "_", true, "▒", true],
		[true, "_", true, "^", true, "▒", true],

		[true, "^", true, "_", true, true],
		[true, "_", true, "^", true, true],

		[true, "┬", true, true],
		[true, "┴", true, true],

		["⏟", true, "┬", true],
		["⏟", true, "┴", true],
		["⏞", true, "┬", true],
		["⏞", true, "┴", true],


		[true, "_", true, "^", true],
		[true, "^", true, "_", true],

		[true, "^", true, "▒", true],
		[true, "_", true, "▒", true],
		[true, "▒", true],

		[true, "┬", true],
		[true, "┴", true],

		[true, "⏟", true],
		[true, "⏞", true],

		["s","i","n"],["t","a","n"],["t","a","n","h"],["s","u","p"],["s","i","n","h"],["s","e","c"],
		["h","o","m"],["a","r","g"],["a","r","c","y","a","n"],["a","r","c","s","i","n"],["a","r","c","s","e","c"],
		["a","r","c","c","s","c"],["a","r","c","c","o","t"],["a","r","c","c","o","s"],["i","n","f"],["g","c","d"],
		["e","x","p"],["d","i","m"],["d","e","t"],["d","e","g"],["c","s","c"],["c","o","t","h"],["c","o","t"],
		["c","o","s","h"],["c","o","s"],["P","r"],["l","g"],["l","n"],["l","o","g"],["s","g","n"],["s","e","c","h"],
		["l","i","m"],["m","i","n"],["m","a","x"],

		[true, "/", true],
		[true, "^", true],
		[true, "_", true],

		['\\frac', true, true],

		[true, "'", true ],
		[true, "′", true ],
		[true, "″", true ],
		[true, "‴", true ],
		[true, "⁗", true ],

		["√", true],
		["∛", true],
		["∜", true],
		["▭", true],
		["□", true],
		["¯", true],
		["▁", true],
		["/", true],
		["⏟", true],
		["⏞", true],
		[true, "/"],

		[true, "┬"],
		[true, "┴"],
		["/"],

		[true, "́" ],
		[true, "̂" ],
		[true, "̌" ],
		[true, "̃" ],
		[true, "̀" ],
		[true, "̇" ],
		[true, "̈" ],
		[true, "⃛" ],
		[true, "̄" ],
		[true, "⃗" ],

		[true, "′" ],
		[true, "'" ],
		[true, "″" ],
		[true, "‴" ],
		[true, "⁗" ],

		["'"],
		["′"],
		["″"],
		["‴"],
		["⁗"],
		[true],
	];

	for (let j = 0; j < rules.length; j++)
	{
		this.cursor = this.Content.length - 1;
		let arrCurrentRule = rules[j];
		let intRuleCounter = 0;

		for (let i = arrCurrentRule.length - 1; i >= 0 && this.IsHasContent(); i--)
		{
			this.CurrentRule = arrCurrentRule[i];
			const intCharCode = this.GetNext(arrCurrentRule[i - 1]);

			if (arrCurrentRule[i] === true && intCharCode)
				intRuleCounter++;
			else if (arrCurrentRule[i] !== true && String.fromCharCode(intCharCode) === arrCurrentRule[i])
				intRuleCounter++;
			else
				break;
		}

		if (intRuleCounter === arrCurrentRule.length)
		{
			this.RulePosition[0] = this.cursor;
			this.RulePosition[1] = this.CurrentElement instanceof ParaRunIterator && this.CurrentElement.Cursor >= 0 ? this.CurrentElement.Cursor + 1: 0;
			break;
		}
		else
			this.ResetParaRunCursor();
	}

	return;
};
ContentIterator.prototype.WriteBracketPosition = function (intGlobalPos, intParaRunPos)
{
    this.BracketPosition[0] = intGlobalPos;
    this.BracketPosition[1] = intParaRunPos;
};
ContentIterator.prototype.GetPosition = function()
{
    let brack = this.BracketPosition;
    let rule = this.RulePosition;

    if (rule[0] > brack[0])
    {
        if (rule[1] > brack[1]) {
            return rule;
        }
        else {
            this.isBracket = true;
            return brack;
        }
    }
    else
    {
        if (brack[1] > rule[1]) {
            this.isBracket = true
            return brack
        }
        else {
            return rule
        }
    }
};
CMathContent.prototype.CheckAutoCorrectionRules = function(nInputType)
{
    const oRuleIterator = new ContentIterator(this);
    let prev = [];
    let isSpace = false;
    let lastOperator;

    if (this.IsLastElement(AscMath.MathLiterals.operators))
    {
        lastOperator = this.GetLastTextElement();
        let lastContent = this.Content[this.Content.length - 1];
        if (lastContent && lastContent.Content.length >= 1)
        {
            lastContent.Remove_FromContent(lastContent.Content.length - 1, 1);
        }
    }
    else if (this.GetLastTextElement() === " ")
    {
        isSpace = this.DeleteEndSpace();
        for (let i = 0; i < this.Content.length; i++) {
            prev.push([this.Content[i].constructor.name, this.Content[i].Content ? this.Content[i].Content.length : 0]);
        }

		// if we have only space - doesn't convert
		if (prev.length === 1 && prev[0][0] === 'ParaRun' && this.Content[0].GetTextOfElement().length === 0)
		{
			this.Add_Text(' ');
			return;
		}
    }

    oRuleIterator.CheckRules();

    const arrPosition = oRuleIterator.RulePosition;

    if (arrPosition.length === 2)
        this.CutConvertAndPaste(arrPosition, nInputType);
    else if (isSpace)
        this.Add_TextOnPos(this.Content.length - 1,' ');

    let now = [];

    for (let i = 0; i < this.Content.length; i++) {
        now.push([this.Content[i].constructor.name, this.Content[i].Content ? this.Content[i].Content.length : 0]);
    }

    if (isSpace)
    {
        let counter = 0;
        let isEqual = true;
        while (counter !== 2)
        {
            let tprev = prev[prev.length - 1 - counter];
            let tnow = now[now.length - 1 - counter];

            let tprevType = tprev ? tprev[0] : undefined;
            let tnowType = tprev ? tnow[0] : undefined;
            let tprevCount = tprev ? tprev[1] : undefined;
            let tnowCount = tprev ? tnow[1] : undefined;

            if (tprevType !== tnowType || tprevCount !== tnowCount)
			{
                if (!(counter === 0 && tprevCount === tnowCount + 1)) {
                    isEqual = false;
                    break;
                }
            }

            if (tnow !== "ParaRun" && counter > 0)
                break;

            counter++;
        }
        if (isEqual)
			this.Add_Text(' ');
    }
    else if (lastOperator)
    {
        this.Add_TextOnPos(this.Content.length, lastOperator);
    }
};
CMathContent.prototype.IsLastTextElementRBracket = function()
{
    let strLast = this.GetLastTextElement()
    return  AscMath.MathLiterals.rBrackets.IsIncludes(strLast) ||
            AscMath.MathLiterals.lrBrackets.IsIncludes(strLast)
};
CMathContent.prototype.IsPreLastTextElementRBracket = function()
{
    let strPreLast = this.GetPreLastTextElement()
    return  AscMath.MathLiterals.rBrackets.IsIncludes(strPreLast) ||
            AscMath.MathLiterals.lrBrackets.IsIncludes(strPreLast)
};
CMathContent.prototype.CheckAutoCorrectionBrackets = function(nInputType)
{
    const oBracketsContent = this.GetBracketOperatorInfo(nInputType === 1);
    const Brackets = new ProceedBrackets(oBracketsContent, nInputType);
    const arrPosition = Brackets.GetPosition();

    if (!(this.IsLastTextElementRBracket() || this.IsPreLastTextElementRBracket()) || this.GetPreLastTextElement() === "〗")
        return Brackets;

    if (Brackets.BracketsPair.length > 0 && this.GetLastTextElement() === " " ||
                                            this.IsLastTextElementRBracket() ||
                                            this.IsLastElement(AscMath.MathLiterals.operators)
    )
    {
        if (arrPosition.length === 2 && Brackets.intCounter === 0)
        {
			arrPosition[1] += 1;
            this.CutConvertAndPaste(arrPosition, nInputType, true);
            Brackets.isConvert = true;
            if (this.GetLastTextElement() === " ")
                this.DeleteEndSpace();
        }
    }

    return Brackets;
};
CMathContent.prototype.CutConvertAndPaste = function(arrPos, nInputType, isNotWrap)
{
    if (arrPos.length === 0)
        arrPos = [0, 0];

    if (arrPos[0] < 0)
        arrPos[0] = 0;

    let strContent = "";

    for (let i = this.Content.length - 1; i >= arrPos[0]; i--)
    {
        const CurrentContent = this.Content[i];
        if (CurrentContent === undefined)
            continue;

        if (i === arrPos[0] && CurrentContent instanceof ParaRun)
        {
            for (let j = CurrentContent.Content.length - 1; j >= arrPos[1]; j--)
            {
                strContent = CurrentContent.Content[j].GetTextOfElement(nInputType === 1) + strContent;
                CurrentContent.Remove_FromContent(j, 1, true);
            }
        }
        else
        {
            if (CurrentContent.Type !== 49)
			{
				if (CurrentContent instanceof CDelimiter)
				{
					strContent = CurrentContent.GetTextOfElement(nInputType === 1).trim() + strContent;
				}
				else
				{
					strContent = "〖" + CurrentContent.GetTextOfElement(nInputType === 1).trim() + "〗" + strContent;
				}
                strContent = strContent.trim();
            }
            else
			{
                strContent = CurrentContent.GetTextOfElement(nInputType === 1).trim() + strContent;
                strContent = strContent.trim();
            }
            this.Remove_FromContent(i, 1, true);
        }
    }

    AscMath.GetConvertContent(nInputType, strContent, this);
    this.Correct_Content(true);
    this.MoveCursorToEndPos();
};
CMathContent.prototype.DeleteEndSpace = function()
{
    const oLastContent = this.Content[this.Content.length - 1];

    if (oLastContent instanceof ParaRun)
    {
        const isReturn = oLastContent.MathAutoCorrection_DeleteLastSpace();
        if (isReturn)
            return oLastContent;

        return false;
    }
};
CMathContent.prototype.IsStartAutoCorrection = function(nInputType, intCode)
{
    return  AscMath.IsStartAutoCorrection(nInputType, intCode)
	    && !(this.IsLastElement(AscMath.MathLiterals.radical) || this.IsLastElement(AscMath.MathLiterals.nary))
};
CMathContent.prototype.Clear_ContentChanges = function()
{
	this.m_oContentChanges.Clear();
};
CMathContent.prototype.Add_ContentChanges = function(Changes)
{
	this.m_oContentChanges.Add(Changes);
};
CMathContent.prototype.Refresh_ContentChanges = function()
{
	this.m_oContentChanges.Refresh();
};
CMathContent.prototype.IsOneElementInContentForGetText = function()
{
    let count = this.Content.length;

    if (count === 2 && this.Content[this.Content.length - 1].Content.length === 0)
    {
        count--
    }
    if (count === 1 && this.Content[0].Content.length > 1)
    {
        count++
    }
    
    return count === 1;
};
CMathContent.prototype.IsFirstLetterIsBracket = function (isLaTeX)
{
    let firstElement = this.Content[this.Content.length - 1];

    if (isLaTeX && firstElement.GetTextOfElement(isLaTeX)[0] === "{")
    {
        return true;
    }
    else if (!isLaTeX && firstElement.GetTextOfElement(isLaTeX)[0] === "(")
    {
        return true;
    }

    return false;
};
CMathContent.prototype.GetMultipleContentForGetText = function(isLaTeX, isNotBrackets, isMustBeBracketsInLaTeX, isMustBeBracketsInUnicode)
{
    let str = "";

    if ((isMustBeBracketsInLaTeX && isLaTeX) || (isMustBeBracketsInUnicode && !isLaTeX))
    {
        if (isMustBeBracketsInLaTeX === true && isLaTeX)
        {
            str = this.GetTextOfElement(isLaTeX);
            if (str.length > 0 && str[0] !== "{")
            {
                str = "{" + this.GetTextOfElement(isLaTeX) + "}";
            }
        }

        if (isMustBeBracketsInUnicode === true && !isLaTeX)
        {
            if (str.length > 0 && str[0] !== "(")
            {
                str =  "(" + this.GetTextOfElement(isLaTeX) + ")";
            }
        }

        return str;
    }

    if (this.IsOneElementInContentForGetText() || this.IsFirstLetterIsBracket(isLaTeX))
    {
        str = this.GetTextOfElement(isLaTeX)
    }
    else
    {
        if (isNotBrackets)
        {
            str = this.GetTextOfElement(isLaTeX)
        }
        else
        {
            str = this.GetTextOfElement(isLaTeX);
            if (!AscMath.functionNames.includes(str) && !(str[0] === "\"" && str[str.length-1] === "\"") && str[0] !== "(")
            {
                str = (isLaTeX === true)
                    ?  "{" + this.GetTextOfElement(isLaTeX) + "}"
                    :  "(" + this.GetTextOfElement(isLaTeX) + ")";
            }
        }
    }

    if (!isLaTeX && isNotBrackets && str[0] === "├" && str[str.length - 1] === "┤")
    {
        return str.slice(1,str.length - 1);
    }

    return this.CheckIsEmpty(str);
};
CMathContent.prototype.haveMixedContent = function()
{
	let isOperator = 0;
	let isNormalText = 0;
	let isCustomContent = 0;

	for (let i = 0; i < this.Content.length; i++)
	{
		let oCurrentContent = this.Content[i];
		if (oCurrentContent instanceof ParaRun)
		{
			if (oCurrentContent.IsContainMathOperators())
				isOperator = 1;
			
			if (!oCurrentContent.IsContainMathOperators())
				isNormalText = 1;
		}
		else
		{
			if (isCustomContent
				|| oCurrentContent instanceof CFraction
				|| oCurrentContent instanceof CDegree
				|| oCurrentContent instanceof CDegreeSubSup)
				return true;
			
			isCustomContent = 1;
		}
		
		if (isOperator + isNormalText + isCustomContent > 1)
			return true;
	}

	return false;
};
CMathContent.prototype.CheckIsEmpty = function(strAtom)
{
    if (strAtom === '⬚')
        return "";
    else
        return strAtom;
};
CMathContent.prototype.GetTextOfElement = function(isLaTeX)
{
	let str = "";

	for (let i = 0; i < this.Content.length; i++)
    {
		str += this.Content[i].GetTextOfElement(isLaTeX);
	}

	return str;
};
CMathContent.prototype.GetTextContent = function(bSelectedText, isLaTeX)
{
	if (undefined === isLaTeX || null === isLaTeX)
		isLaTeX = false;

	let str = "";
	let StartPos = 0;
	let EndPos = this.Content.length;

	if (bSelectedText) {
		StartPos = (this.Selection.Use == true ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos);
		EndPos   = (this.Selection.Use == true ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos);
	}

	for (let i = StartPos; i <= EndPos; i++)
	{
		if (this.Content[i] !== undefined)
		{
			str += this.Content[i].GetTextOfElement(isLaTeX);
		}
	}

	return {str: str};
};

var g_DefaultAutoCorrectMathFuncs =
[
    'arcsin', 'asin', 'sin', 'arcsinh', 'asinh', 'sinh',
    'arcsec', 'sec', 'asec', 'arcsech', 'asech','sech',
    'arccos', 'acos', 'cos', 'arccosh','acosh', 'cosh',
    'arccsc', 'acsc', 'csc', 'arccsch', 'acsch', 'csch',
    'arctan', 'atan', 'tan', 'arctanh', 'atanh', 'tanh',
    'arccot', 'acot', 'cot', 'arccoth', 'acoth', 'coth',
    'arg', 'det', 'exp', 'inf', 'lim', 'min',
    'def', 'dim', 'gcd', 'ker', 'log', 'Pr',
    'deg', 'erf', 'hom', 'lg', 'ln', 'max', 'sup'
];
var g_DefaultAutoCorrectMathSymbolsList =
[
    ['!!', 0x203C],
    ['...', 0x2026],
    ['::', 0x2237],
    [':=', 0x2254],
    ['/<', 0x226E],
    ['/>', 0x226F],
    ['/=', 0x2260],
    ['~=', 0x2245],
    ['-+', 0x2213],
    ['+-', 0x00B1],
    ['<<', 0x226A],
    ['<=', 0x2264],
    ['->', 0x2192],
    ['>=', 0x2265],
    ['>>', 0x226B],
    ['\\above', 0x2534],
    ['\\acute', 0x0301],
    ['\\aleph', 0x2135],
    ['\\alpha', 0x03B1],
    ['\\Alpha', 0x0391],
    ['\\amalg', 0x2210],
    ['\\angle', 0x2220],
    ['\\aoint', 0x222E],
    ['/\\approx', 0x2249],
    ['\\approx', 0x2248],
    ['\\asmash', 0x2B06],
    ['\\ast', 0x2217],
    ['/\\asymp', 0x226D],
    ['\\asymp', 0x224D],
    ['\\atop', 0x00A6],
    ['\\bar', 0x0305],
    ['\\Bar', 0x033F],
    ['\\because', 0x2235],
    ['\\begin', 0x3016],
    ['\\below', 0x252C],
    ['\\bet', 0x2136],
    ['\\beta', 0x03B2],
    ['\\Beta', 0x0392],
    ['\\beth', 0x2136],
    ['\\bigcap', 0x22C2],
    ['\\bigcup', 0x22C3],
    ['\\bigodot', 0x2A00],
    ['\\bigoplus', 0x2A01],
    ['\\bigotimes', 0x2A02],
    ['\\bigsqcup', 0x2A06],
    ['\\biguplus', 0x2A04],
    ['\\bigvee', 0x22C1],
    ['\\bigwedge', 0x22C0],
    ['\\binomial',[0x0028, 0x0061, 0x002B, 0x0062, 0x0029, 0x005E, 0x005E, 0x003D, 0x2211, 0x005F, 0x0028, 0x006B, 0x003D, 0x0030, 0x0029, 0x005E, 0x006E, 0x0020, 0x2592, 0x0028, 0x006E, 0x00A6, 0x006B, 0x0029, 0x0061, 0x005E, 0x006B, 0x0020, 0x0062, 0x005E, 0x0028, 0x006E, 0x002D, 0x006B, 0x0029]],
    ['\\bot', 0x22A5],
    ['\\bowtie', 0x22C8],
    ['\\box', 0x25A1],
    ['\\boxdot', 0x22A1],
    ['\\boxminus', 0x229F],
    ['\\boxplus', 0x229E],
    ['\\bra', 0x27E8],
    ['\\break', 0x2936],
    ['\\breve', 0x0306],
    ['\\bullet', 0x2219],
    ['\\cap', 0x2229],
    ['\\cbrt', 0x221B],
    ['\\cases', 0x24B8],
    ['\\cdot', 0x22C5],
    ['\\cdots', 0x22EF],
    ['\\check', 0x030C],
    ['\\chi', 0x03C7],
    ['\\Chi', 0x03A7],
    ['\\circ', 0x2218],
    ['\\close', 0x2524],
    ['\\clubsuit', 0x2663],
    ['\\coint', 0x2232],
    ['/\\cong', 0x2247],
    ['\\cong', 0x2245],
    ['\\coprod', 0x2210],
    ['\\cup', 0x222A],
    ['\\dalet', 0x2138],
    ['\\daleth', 0x2138],
    ['\\dashv', 0x22A3],
    ['\\dd', 0x2146],
    ['\\Dd', 0x2145],
    ['\\ddddot', 0x20DC],
    ['\\dddot', 0x20DB],
    ['\\ddot', 0x0308],
    ['\\ddots', 0x22F1],
    ['\\defeq', 0x225D],
    ['\\degc', 0x2103],
    ['\\degf', 0x2109],
    ['\\degree', 0x00B0],
    ['\\delta', 0x03B4],
    ['\\Delta', 0x0394],
    ['\\Deltaeq', 0x225C],
    ['\\diamond', 0x22C4],
    ['\\diamondsuit', 0x2662],
    ['\\div', 0x00F7],
    ['\\dot', 0x0307],
    ['\\doteq', 0x2250],
    ['\\dots', 0x2026],
    ['\\doublea', 0x1D552],
    ['\\doubleA', 0x1D538],
    ['\\doubleb', 0x1D553],
    ['\\doubleB', 0x1D539],
    ['\\doublec', 0x1D554],
    ['\\doubleC', 0x2102],
    ['\\doubled', 0x1D555],
    ['\\doubleD', 0x1D53B],
    ['\\doublee', 0x1D556],
    ['\\doubleE', 0x1D53C],
    ['\\doublef', 0x1D557],
    ['\\doubleF', 0x1D53D],
    ['\\doubleg', 0x1D558],
    ['\\doubleG', 0x1D53E],
    ['\\doubleh', 0x1D559],
    ['\\doubleH', 0x210D],
    ['\\doublei', 0x1D55A],
    ['\\doubleI', 0x1D540],
    ['\\doublej', 0x1D55B],
    ['\\doubleJ', 0x1D541],
    ['\\doublek', 0x1D55C],
    ['\\doubleK', 0x1D542],
    ['\\doublel', 0x1D55D],
    ['\\doubleL', 0x1D543],
    ['\\doublem', 0x1D55E],
    ['\\doubleM', 0x1D544],
    ['\\doublen', 0x1D55F],
    ['\\doubleN', 0x2115],
    ['\\doubleo', 0x1D560],
    ['\\doubleO', 0x1D546],
    ['\\doublep', 0x1D561],
    ['\\doubleP', 0x2119],
    ['\\doubleq', 0x1D562],
    ['\\doubleQ', 0x211A],
    ['\\doubler', 0x1D563],
    ['\\doubleR', 0x211D],
    ['\\doubles', 0x1D564],
    ['\\doubleS', 0x1D54A],
    ['\\doublet', 0x1D565],
    ['\\doubleT', 0x1D54B],
    ['\\doubleu', 0x1D566],
    ['\\doubleU', 0x1D54C],
    ['\\doublev', 0x1D567],
    ['\\doubleV', 0x1D54D],
    ['\\doublew', 0x1D568],
    ['\\doubleW', 0x1D54E],
    ['\\doublex', 0x1D569],
    ['\\doubleX', 0x1D54F],
    ['\\doubley', 0x1D56A],
    ['\\doubleY', 0x1D550],
    ['\\doublez', 0x1D56B],
    ['\\doubleZ', 0x2124],
    ['\\downarrow', 0x2193],
    ['\\Downarrow', 0x21D3],
    ['\\dsmash', 0x2B07],
    ['\\ee', 0x2147],
    ['\\ell', 0x2113],
    ['\\emptyset', 0x2205],
    ['\\emsp', 0x2003],
    ['\\end', 0x3017],
    ['\\ensp', 0x2002],
    ['\\epsilon', 0x03F5],
    ['\\Epsilon', 0x0395],
    ['\\eqarray', 0x2588],
    ['/\\equiv', 0x2262],
    ['\\equiv', 0x2261],
    ['\\eta', 0x03B7],
    ['\\Eta', 0x0397],
    ['/\\exists', 0x2204],
    ['\\exists', 0x2203],
    ['\\forall', 0x2200],
    ['\\fraktura', 0x1D51E],
    ['\\frakturA', 0x1D504],
    ['\\frakturb', 0x1D51F],
    ['\\frakturB', 0x1D505],
    ['\\frakturc', 0x1D520],
    ['\\frakturC', 0x212D],
    ['\\frakturd', 0x1D521],
    ['\\frakturD', 0x1D507],
    ['\\frakture', 0x1D522],
    ['\\frakturE', 0x1D508],
    ['\\frakturf', 0x1D523],
    ['\\frakturF', 0x1D509],
    ['\\frakturg', 0x1D524],
    ['\\frakturG', 0x1D50A],
    ['\\frakturh', 0x1D525],
    ['\\frakturH', 0x210C],
    ['\\frakturi', 0x1D526],
    ['\\frakturI', 0x2111],
    ['\\frakturj', 0x1D527],
    ['\\frakturJ', 0x1D50D],
    ['\\frakturk', 0x1D528],
    ['\\frakturK', 0x1D50E],
    ['\\frakturl', 0x1D529],
    ['\\frakturL', 0x1D50F],
    ['\\frakturm', 0x1D52A],
    ['\\frakturM', 0x1D510],
    ['\\frakturn', 0x1D52B],
    ['\\frakturN', 0x1D511],
    ['\\frakturo', 0x1D52C],
    ['\\frakturO', 0x1D512],
    ['\\frakturp', 0x1D52D],
    ['\\frakturP', 0x1D513],
    ['\\frakturq', 0x1D52E],
    ['\\frakturQ', 0x1D514],
    ['\\frakturr', 0x1D52F],
    ['\\frakturR', 0x211C],
    ['\\frakturs', 0x1D530],
    ['\\frakturS', 0x1D516],
    ['\\frakturt', 0x1D531],
    ['\\frakturT', 0x1D517],
    ['\\frakturu', 0x1D532],
    ['\\frakturU', 0x1D518],
    ['\\frakturv', 0x1D533],
    ['\\frakturV', 0x1D519],
    ['\\frakturw', 0x1D534],
    ['\\frakturW', 0x1D51A],
    ['\\frakturx', 0x1D535],
    ['\\frakturX', 0x1D51B],
    ['\\fraktury', 0x1D536],
    ['\\frakturY', 0x1D51C],
    ['\\frakturz', 0x1D537],
    ['\\frakturZ', 0x2128],
    ['\\frown', 0x2311],
    ['\\funcapply', 0x2061],
    ['\\G', 0x0393],
    ['\\gamma', 0x03B3],
    ['\\Gamma', 0x0393],
    ['/\\ge', 0x2271],
    ['\\ge', 0x2265],
    ['\\geq', 0x2265],
    ['\\gets', 0x2190],
    ['/\\gtrless', 0x2279],
    ['\\gtrless', 0x2277],
    ['\\gg', 0x226B],
    ['\\gimel', 0x2137],
    ['\\grave', 0x0300],
    ['\\hairsp', 0x200A],
    ['\\hat', 0x0302],
    ['\\hbar', 0x210F],
    ['\\heartsuit', 0x2661],
    ['\\hookleftarrow', 0x21A9],
    ['\\hookrightarrow', 0x21AA],
    ['\\hphantom', 0x2B04],
    ['\\hsmash', 0x2B0C],
    ['\\hvec', 0x20D1],
    ['\\identitymatrix', [0x0028, 0x25A0, 0x0028, 0x0031, 0x0026, 0x0030, 0x0026, 0x0030, 0x0040, 0x0030, 0x0026, 0x0031, 0x0026, 0x0030, 0x0040, 0x0030, 0x0026, 0x0030, 0x0026, 0x0031, 0x0029, 0x0029]],
    ['\\ii', 0x2148],
    ['\\iiint', 0x222D],
    ['\\iint', 0x222C],
    ['\\iiiint', 0x2A0C],
    ['\\Im', 0x2111],
    ['\\imath', 0x0131],
    ['/\\in', 0x2209],
    ['\\in', 0x2208],
    ['\\inc', 0x2206],
    ['\\infty', 0x221E],
    ['\\int', 0x222B],
    ['\\integral', [0x0031, 0x002F, 0x0032, 0x03C0, 0x222B, 0x005F, 0x0030, 0x005E, 0x0032, 0x03C0, 0x2592, 0x2146, 0x03B8, 0x0020, 0x0028, 0x0061, 0x002B, 0x0062, 0x0073, 0x0069, 0x006E, 0x0020, 0x03B8, 0x0029, 0x003D, 0x0031, 0x002F, 0x221A, 0x0028, 0x0061, 0x005E, 0x0032, 0x002D, 0x0062, 0x005E, 0x0032, 0x0029]],
    ['\\iota', 0x03B9],
    ['\\Iota', 0x0399],
    ['\\itimes', 0x2062],
    ['\\j', [0x004A, 0x0061, 0x0079]],
    ['\\jj', 0x2149],
    ['\\jmath', 0x0237],
    ['\\kappa', 0x03BA],
    ['\\Kappa', 0x039A],
    ['\\ket', 0x27E9],
    ['\\lambda', 0x03BB],
    ['\\Lambda', 0x039B],
    ['\\langle', 0x2329],
    ['\\lbbrack', 0x27E6],
    ['\\lbrace', 0x007B],
    ['\\lbrack', 0x005B],
    ['\\lceil', 0x2308],
    ['\\ldiv', 0x2215],
    ['\\ldivide', 0x2215],
    ['\\ldots', 0x2026],
    ['/\\le', 0x2270],
    ['\\le', 0x2264],
    ['\\left', 0x251C],
    ['\\leftarrow', 0x2190],
    ['\\Leftarrow', 0x21D0],
    ['\\leftharpoondown', 0x21BD],
    ['\\leftharpoonup', 0x21BC],
    ['\\leftrightarrow', 0x2194],
    ['\\Leftrightarrow', 0x21D4],
    ['\\leq', 0x2264],
    ['/\\lessgtr', 0x2278],
    ['\\lessgtr', 0x2276],
    ['\\lfloor', 0x230A],
    ['\\lhvec', 0x20D0],
    ['\\limit', [0x006C, 0x0069, 0x006D, 0x005F, 0x0028, 0x006E, 0x2192, 0x221E, 0x0029, 0x2061, 0x3016, 0x0028, 0x0031, 0x002B, 0x0031, 0x002F, 0x006E, 0x0029, 0x005E, 0x006E, 0x3017, 0x003D, 0x0065]],
    ['\\ll', 0x226A],
    ['\\lmoust', 0x23B0],
    ['\\Longleftarrow', 0x27F8],
    ['\\Longleftrightarrow', 0x27FA],
    ['\\Longrightarrow', 0x27F9],
    ['\\lrhar', 0x21CB],
    ['\\lvec', 0x20D6],
    ['\\mapsto', 0x21A6],
    ['\\matrix', 0x25A0],
    ['\\medsp', 0x205F],
    ['\\mid', 0x2223],
    ['\\middle', 0x24DC],
    ['\\models', 0x22A8],
    ['\\mp', 0x2213],
    ['\\mu', 0x03BC],
    ['\\Mu', 0x039C],
    ['\\nabla', 0x2207],
    ['\\naryand', 0x2592],
    ['\\nbsp', 0x00A0],
    ['\\ne', 0x2260],
    ['\\nearrow', 0x2197],
    ['\\neq', 0x2260],
    ['/\\ni', 0x220C],
    ['\\ni', 0x220B],
    ['\\norm', 0x2016],
    ['\\notcontain', 0x220C],
    ['\\notelement', 0x2209],
    ['\\notin', 0x2209],
    ['\\nu', 0x03BD],
    ['\\Nu', 0x039D],
    ['\\nwarrow', 0x2196],
    ['\\o', 0x03BF],
    ['\\O', 0x039F],
    ['\\odot', 0x2299],
    ['\\of', 0x2592],
    ['\\oiiint', 0x2230],
    ['\\oiint', 0x222F],
    ['\\oint', 0x222E],
    ['\\omega', 0x03C9],
    ['\\Omega', 0x03A9],
    ['\\ominus', 0x2296],
    ['\\open', 0x251C],
    ['\\oplus', 0x2295],
    ['\\otimes', 0x2297],
    ['\\over', 0x002F],
    ['\\overbar', 0x00AF],
    ['\\overbrace', 0x23DE],
    ['\\overbracket', 0x23B4],
    ['\\overline', 0x00AF],
    ['\\overparen', 0x23DC],
    ['\\overshell', 0x23E0],
    ['\\parallel', 0x2225],
    ['\\partial', 0x2202],
    ['\\pmatrix', 0x24A8],
    ['\\perp', 0x22A5],
    ['\\phantom', 0x27E1],
    ['\\phi', 0x03D5],
    ['\\Phi', 0x03A6],
    ['\\pi', 0x03C0],
    ['\\Pi', 0x03A0],
    ['\\pm', 0x00B1],
    ['\\pppprime', 0x2057],
    ['\\ppprime', 0x2034],
    ['\\pprime', 0x2033],
    ['/\\prec', 0x2280],
    ['\\prec', 0x227A],
    ['/\\preceq', 0x22E0],
    ['\\preceq', 0x227C],
    ['\\prime', 0x2032],
    ['\\prod', 0x220F],
    ['\\propto', 0x221D],
    ['\\psi', 0x03C8],
    ['\\Psi', 0x03A8],
    ['\\qdrt', 0x221C],
    ['\\quadratic', [0x0078, 0x003d, 0x0028, 0x002d, 0x0062, 0x00B1, 0x221A, 0x0028, 0x0062, 0x005e, 0x0032, 0x002d, 0x0034, 0x0061, 0x0063, 0x0029, 0x0029, 0x002f, 0x0032, 0x0061]],
    ['\\rangle', 0x232A],
    ['\\Rangle', 0x27EB],
    ['\\ratio', 0x2236],
    ['\\rbrace', 0x007D],
    ['\\rbrack', 0x005D],
    ['\\Rbrack', 0x27E7],
    ['\\rceil', 0x2309],
    ['\\rddots', 0x22F0],
    ['\\Re', 0x211C],
    ['\\rect', 0x25AD],
    ['\\rfloor', 0x230B],
    ['\\rho', 0x03C1],
    ['\\Rho', 0x03A1],
    ['\\rhvec', 0x20D1],
    ['\\right', 0x2524],
    ['\\rightarrow', 0x2192],
    ['\\Rightarrow', 0x21D2],
    ['\\rightharpoondown', 0x21C1],
    ['\\rightharpoonup', 0x21C0],
    ['\\rmoust', 0x23B1],
    ['\\root', 0x24AD],
    ['\\scripta', 0x1D4B6],
    ['\\scriptA', 0x1D49C],
    ['\\scriptb', 0x1D4B7],
    ['\\scriptB', 0x212C],
    ['\\scriptc', 0x1D4B8],
    ['\\scriptC', 0x1D49E],
    ['\\scriptd', 0x1D4B9],
    ['\\scriptD', 0x1D49F],
    ['\\scripte', 0x212F],
    ['\\scriptE', 0x2130],
    ['\\scriptf', 0x1D4BB],
    ['\\scriptF', 0x2131],
    ['\\scriptg', 0x210A],
    ['\\scriptG', 0x1D4A2],
    ['\\scripth', 0x1D4BD],
    ['\\scriptH', 0x210B],
    ['\\scripti', 0x1D4BE],
    ['\\scriptI', 0x2110],
    ['\\scriptj', 0x1D4BF],
    ['\\scriptJ', 0x1D4A5],
    ['\\scriptk', 0x1D4C0],
    ['\\scriptK', 0x1D4A6],
    ['\\scriptl', 0x2113],
    ['\\scriptL', 0x2112],
    ['\\scriptm', 0x1D4C2],
    ['\\scriptM', 0x2133],
    ['\\scriptn', 0x1D4C3],
    ['\\scriptN', 0x1D4A9],
    ['\\scripto', 0x2134],
    ['\\scriptO', 0x1D4AA],
    ['\\scriptp', 0x1D4C5],
    ['\\scriptP', 0x1D4AB],
    ['\\scriptq', 0x1D4C6],
    ['\\scriptQ', 0x1D4AC],
    ['\\scriptr', 0x1D4C7],
    ['\\scriptR', 0x211B],
    ['\\scripts', 0x1D4C8],
    ['\\scriptS', 0x1D4AE],
    ['\\scriptt', 0x1D4C9],
    ['\\scriptT', 0x1D4AF],
    ['\\scriptu', 0x1D4CA],
    ['\\scriptU', 0x1D4B0],
    ['\\scriptv', 0x1D4CB],
    ['\\scriptV', 0x1D4B1],
    ['\\scriptw', 0x1D4CC],
    ['\\scriptW', 0x1D4B2],
    ['\\scriptx', 0x1D4CD],
    ['\\scriptX', 0x1D4B3],
    ['\\scripty', 0x1D4CE],
    ['\\scriptY', 0x1D4B4],
    ['\\scriptz', 0x1D4CF],
    ['\\scriptZ', 0x1D4B5],
    ['\\sdiv', 0x2044],
    ['\\sdivide', 0x2044],
    ['\\searrow', 0x2198],
    ['\\setminus', 0x2216],
    ['\\sigma', 0x03C3],
    ['\\Sigma', 0x03A3],
    ['/\\sim', 0x2241],
    ['\\sim', 0x223C],
    ['/\\simeq', 0x2244],
    ['\\simeq', 0x2243],
    ['\\smash', 0x2B0D],
    ['\\smile', 0x2323],
    ['\\spadesuit', 0x2660],
    ['\\sqcap', 0x2293],
    ['\\sqcup', 0x2294],
    ['\\sqrt', 0x221A],
    ['/\\sqsubseteq', 0x22E2],
    ['\\sqsubseteq', 0x2291],
    ['/\\sqsuperseteq', 0x22E3],
    ['\\sqsuperseteq', 0x2292],
    ['/\\sqsupseteq', 0x22E3],
    ['\\sqsupseteq', 0x2292],
    ['\\star', 0x22C6],
    ['/\\subset', 0x2284],
    ['\\subset', 0x2282],
    ['/\\subseteq', 0x2288],
    ['\\subseteq', 0x2286],
    ['/\\succ', 0x2281],
    ['\\succ', 0x227B],
    ['/\\succeq', 0x22E1],
    ['\\succeq', 0x227D],
    ['\\sum', 0x2211],
    ['/\\supset', 0x2285],
    ['\\supset', 0x2283],
    ['/\\superset', 0x2285],
    ['\\superset', 0x2283],
    ['/\\superseteq', 0x2289],
    ['\\superseteq', 0x2287],
    ['/\\supseteq', 0x2289],
    ['\\supseteq', 0x2287],
    ['\\swarrow', 0x2199],
    ['\\tau', 0x03C4],
    ['\\Tau', 0x03A4],
    ['\\therefore', 0x2234],
    ['\\theta', 0x03B8],
    ['\\Theta', 0x0398],
    ['\\thicksp', 0x2005],
    ['\\thinsp', 0x2006],
    ['\\tilde', 0x0303],
    ['\\times', 0x00D7],
    ['\\to', 0x2192],
    ['\\top', 0x22A4],
    ['\\tvec', 0x20E1],
    ['\\ubar', 0x0332],
    ['\\Ubar', 0x0333],
    ['\\underbar', 0x2581],
    ['\\underbrace', 0x23DF],
    ['\\underbracket', 0x23B5],
    ['\\underline', 0x25B1],
    ['\\underparen', 0x23DD],
    ['\\uparrow', 0x2191],
    ['\\Uparrow', 0x21D1],
    ['\\updownarrow', 0x2195],
    ['\\Updownarrow', 0x21D5],
    ['\\uplus', 0x228E],
    ['\\upsilon', 0x03C5],
    ['\\Upsilon', 0x03A5],
    ['\\varepsilon', 0x03B5],
    ['\\varphi', 0x03C6],
    ['\\varpi', 0x03D6],
    ['\\varrho', 0x03F1],
    ['\\varsigma', 0x03C2],
    ['\\vartheta', 0x03D1],
    ['\\vbar', 0x2502],
    ['\\vdash', 0x22A2],
    ['\\vdots', 0x22EE],
    ['\\vec', 0x20D7],
    ['\\vee', 0x2228],
    ['\\vert', 0x007C],
    ['\\Vert', 0x2016],
    ['\\Vmatrix', 0x24A9],
    ['\\vphantom', 0x21F3],
    ['\\vthicksp', 0x2004],
    ['\\wedge', 0x2227],
    ['\\wp', 0x2118],
    ['\\wr', 0x2240],
    ['\\xi', 0x03BE],
    ['\\Xi', 0x039E],
    ['\\zeta', 0x03B6],
    ['\\Zeta', 0x0396],
    ['\\zwnj', 0x200C],
    ['\\zwsp', 0x200B]
];
var g_AutoCorrectMathSymbols = JSON.parse(JSON.stringify(g_DefaultAutoCorrectMathSymbolsList));
var g_AutoCorrectMathFuncs = JSON.parse(JSON.stringify(g_DefaultAutoCorrectMathFuncs));
var g_AutoCorrectMathsList = {
    DefaultAutoCorrectMathSymbolsList : g_DefaultAutoCorrectMathSymbolsList,
    AutoCorrectMathSymbols : g_AutoCorrectMathSymbols,
    DefaultAutoCorrectMathFuncs : g_DefaultAutoCorrectMathFuncs,
    AutoCorrectMathFuncs : g_AutoCorrectMathFuncs
};
//символы для mathfunc (интеграл, сумма...)
var q_aMathAutoCorrectControlAggregationCodes =
{
    0x2211 : 1, 0x220F : 1, 0x2210 : 1, 0x22C0 : 1, 0x222E : 1,
    0x22C1 : 1, 0x22C2 : 1, 0x22C3 : 1, 0x2A06 : 1, 0x2A04 : 1,
    0x2A00 : 1, 0x2A01 : 1, 0x2A02 : 1, 0x222B : 1, 0x222C : 1,
    0x222D : 1, 0x2A0C : 1, 0x222F : 1, 0x2230 : 1, 0x2232 : 1
};
//символы accent
var q_aMathAutoCorrectAccentCharCodes =
{
    0x0305 : 1, 0x033F : 1, 0x0332 : 1, 0x0333 : 1, 0x0301 : 1,
    0x0300 : 1, 0x20D7 : 1, 0x0306 : 1, 0x0302 : 1, 0x20E1 : 1,
    0x20D1 : 1, 0x030C : 1, 0x0303 : 1, 0x20D6 : 1, 0x20D0 : 1,
    0x0307 : 1, 0x0308 : 1, 0x20DB : 1, 0x20DC : 1, 0x2032 : 1,
    0x2033 : 1, 0x2034 : 1, 0x2057 : 1
};
//left brackets
var g_MathLeftBracketAutoCorrectCharCodes =
{
    0x28 : 1, 0x5B : 1, 0x7B : 1, /*0x7C : 1, 0x2016 : 1,*/ 0x27E8 : 1,
    0x2329 : 1, 0x27E6 : 1, 0x27EA : 1, 0x2308 : 1, 0x230A : 1,
    0x3016 : 1, 0x251C : 1
};
//righ brackets
var g_MathRightBracketAutoCorrectCharCodes =
{
    0x29 : 1, 0x5D : 1, 0x7D : 1, /*0x7C : 1, 0x2016 : 1,*/ 0x27E9 : 1,
    0x232A : 1, 0x27E7 : 1, 0x27EB : 1, 0x2309 : 1, 0x230B : 1,   
    0x3017 : 1, 0x2524 : 1
};
//знаки (минус, сумма...)
var g_aMathAutoCorrectFracCharCodes =
{
    0x20 : 1, 0x21 : 1, /*0x22 : 1,*/ 0x23 : 1,	0x24 : 1, 0x25 : 1, 0x26 : 1,
    /*0x27 : 1,*/ 0x28 : 1, 0x29 : 1, 0x2A : 1, 0x2B : 1, 0x2C : 1, 0x2D : 1,
    0x2E : 1, 0x2F : 1, 0x3A : 1, 0x3B : 1, 0x3C : 1, 0x3D : 1, 0x3E : 1,
    0x3F : 1, 0x40 : 1, 0x5B : 1, /*0x5C : 1,*/ 0x5D : 1, 0x5E : 1, 0x5F : 1,
    0x60 : 1, 0x7B : 1, /*0x7C : 1,*/ 0x7D : 1, 0x7E : 1, /*0x2592 : 1,*/ 0xD7 : 1,
	0x221D: 1, 0x2248: 1, 0x2249: 1, 0x2261: 1, 0x2262: 1, 0x00F7 : 1,
	0x226A: 1, 0x226B: 1, 0x203C: 1, 0x2026: 1, 0x2237: 1, 0x2254: 1, 0x226E: 1,
	0x226F: 1, 0x2260: 1, 0x2245: 1, 0x2213: 1, 0x00B1: 1, 0x2264: 1,
	0x2265: 1, 0x2200: 1, 0x222A: 1, 0x2229: 1, 0x2205: 1, 0x00B0: 1, 0x2109: 1,
	0x2103: 1, 0x2203: 1, 0x2204: 1, 0x2208: 1, 0x220B: 1, 0x2190: 1, 0x2191: 1,
	0x2192: 1, 0x2193: 1, 0x2194: 1, 0x2234: 1, 0x2217: 1, 0x2219: 1, 0x22EE: 1,
	0x22EF: 1, 0x22F0: 1, 0x22F1: 1, 0x2062: 1, 0x2297: 1
};
//символы для определения необходимости автозамены
var g_aMathAutoCorrectTriggerCharCodes =
{
    0x20 : 1, 0x21 : 1, 0x22 : 1, 0x23 : 1, 0x24 : 1, 0x25 : 1, 0x26 : 1,
    0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2A : 1, 0x2B : 1, 0x2C : 1, 0x2D : 1,
    0x2E : 1, 0x2F : 1, 0x3A : 1, 0x3B : 1, 0x3C : 1, 0x3D : 1, 0x3E : 1,
    0x3F : 1, 0x40 : 1, 0x5B : 1, 0x5D : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1,
    0x60 : 1, 0x7B : 1, 0x7D : 1, 0x7C : 1, 0x7E : 1 /*,0x2592 : 1*/
};
//символы при которых производится автозамена текстовой функции (lim, max, sin...)
var g_aMathAutoCorrectTextFunc = {
    0x20 : 1, 0x28 : 1, 0x29 : 1, 0x5B : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1
};
//символы, при которых не производится автозамена дроби без скобок
var g_aMathAutoCorrectNotDoFraction = {
    0x21 : 1, 0x22 : 1, 0x27 : 1, 0x24 : 1, 0x28 : 1, 0x29 : 1, 0x5B : 1,
    0x5D : 1, 0x7B : 1, 0x7D : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1, 0x7C : 1
};
//символы, при которых не производится автозамена степени без скобок
var g_aMathAutoCorrectDoNotDegree = {
    0x21 : 1, 0x22 : 1, 0x27 : 1, 0x24 : 1, 0x28 : 1, 0x29 : 1, 0x5B : 1, 0x5D : 1,
    0x7B : 1, 0x7D : 1, 0x2C : 1, 0x2E : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1, 0x7C : 1
};
//символы, при которых не производится автозамена nary
var g_aMathAutoCorrectNotDoCNary = {
   0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2E : 1, 0x5B : 1,
   0x5C : 1, 0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1, 0x7D : 1, 0x7C : 1 
};
//символы, при которых не производится автозамена скобок
var g_aMathAutoCorrectDoNotDelimiter = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2F : 1, 0x5B : 1,
    0x5C : 1, 0x5E : 1, 0x5F : 1, 0x5D : 1, 0x7B : 1, 0x7D : 1, 0x7C : 1
 };
// символы, при которых не производится автозамена корня
var g_aMathAutoCorrectDoNotRadical = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x5C : 1,
    0x5E : 1, 0x5F : 1, 0x5D : 1, 0x7D : 1
};
// символы, при которых не производится автозамена Box
var g_aMathAutoCorrectDoNotBox = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1,
    0x5B : 1, 0x5D : 1, 0x7B : 1, 0x7D : 1, 0x5C : 1, 0x5E : 1,
    0x5F : 1, 0x7C : 1,
};
// символы, при которых не производится автозамена Group character
var g_aMathAutoCorrectDoNotGroupChar = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2C : 1,
    0x2E : 1, 0x5B : 1, 0x5D : 1, 0x7B : 1, 0x7D : 1, 0x5C : 1, 0x7C : 1
};
// символы, при которых не производится автозамена EqArray и Matrix
var g_aMathAutoCorrectDoNotMatrix = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x5C : 1,
    0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7D : 1
};
// символы, при которых не производится автозамена Accent не вплотную
var g_aMathAutoCorrectDoNotAccentNotClose = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x5C : 1,
    0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7D : 1
};
// символы, при которых не производится автозамена Accent вплотную
var g_aMathAutoCorrectDoNotAccentClose = {
    0x22 : 1, 0x29 : 1, 0x5D : 1, 0x7D : 1
};
// символы при которых не производится автозамена Above
var g_aMathAutoCorrectDoNotAbove = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1,
    0x29 : 1, 0x2C : 1, 0x2E : 1, 0x5B : 1, 0x5C : 1,
    0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1, 0x7C : 1, 0x7D : 1
};
// символы при которых не производится автозамена mathfunc 0x2061
var g_aMathAutoCorrectDoNotMathFunc = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1,
    0x28 : 1, 0x29 : 1, 0x2C : 1, 0x2E : 1,
    0x5B : 1, 0x5C : 1, 0x5D : 1, 0x5E : 1,
    0x5F : 1, 0x7B : 1, 0x7C : 1, 0x7D : 1
};
// символы корней
var g_aMathAutoCorrectRadicalCharCode = {
    0x221A : 1, 0x221B : 1, 0x221C : 1
};
//символы, при которых надо пропускать обработку скобок
//во вложенном контенте
var g_aMathAutoCorrectSkipBrackets = {
    0x221A : 1, 0x25A1 : 1, 0x25AD : 1, 0x23DE : 1, 0x23DC : 1,
    0x23E0 : 1, 0x23DF : 1, 0x23DD : 1, 0x00AF : 1, 0x2581 : 1,
    0x24B8 : 1, 0x2588 : 1, 0x25A0 : 1, 0x24A8 : 1, 0x24A9 : 1,
    0x2534 : 1, 0x252C: 1
};
//символы MATH_GROUP_CHARACTER
var g_aMathAutoCorrectGroupChar = {
    0x23DE : 1, 0x23DC : 1, 0x23E0 : 1, 0x23DF : 1, 0x23DD : 1
};
//symbols matrix and eqarray
var g_aMathAutoCorrectEqArrayMatrix = {
    0x24B8 : 1, 0x2588 : 1, 0x25A0 : 1, 0x24A8 : 1, 0x24A9 : 1
};
// latin alphabet ascii code
var g_aMathAutoCorrectLatinAlph = {
    0x41 : 1, 0x42 : 1, 0x43 : 1, 0x44 : 1, 0x45 : 1, 0x46 : 1,
    0x47 : 1, 0x48 : 1, 0x49 : 1, 0x4A : 1, 0x4B : 1, 0x4C : 1,
    0x4D : 1, 0x4E : 1, 0x4F : 1, 0x50 : 1, 0x51 : 1, 0x52 : 1,
    0x53 : 1, 0x54 : 1, 0x55 : 1, 0x56 : 1, 0x57 : 1, 0x58 : 1,
    0x59 : 1, 0x5A : 1, 0x60 : 1, 0x61 : 1, 0x62 : 1, 0x63 : 1,
    0x64 : 1, 0x65 : 1, 0x66 : 1, 0x67 : 1, 0x68 : 1, 0x69 : 1,
    0x6A : 1, 0x6B : 1, 0x6C : 1, 0x6D : 1, 0x6E : 1, 0x6F : 1,
    0x70 : 1, 0x71 : 1, 0x72 : 1, 0x73 : 1, 0x74 : 1, 0x75 : 1,
    0x76 : 1, 0x77 : 1, 0x78 : 1, 0x79 : 1, 0x7A : 1
};
// special symbols for autocorrect text
var g_aMathAutoCorrectSpecSymb = [
    '!!', '...', '::', ':=', '/<', '/>', '/=',
    '~=', '-+', '+-', '<<', '<=', '->', '>=', '>>'
];

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CMathContent = CMathContent;
window['AscCommonWord'].g_AutoCorrectMathsList = g_AutoCorrectMathsList;

