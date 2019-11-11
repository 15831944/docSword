//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     ANTLR Version: 4.5.1
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// Generated from calcExpr.g4 by ANTLR 4.5.1

// Unreachable code detected
#pragma warning disable 0162
// The variable '...' is assigned but its value is never used
#pragma warning disable 0219
// Missing XML comment for publicly visible type or member '...'
#pragma warning disable 1591

using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using IToken = Antlr4.Runtime.IToken;

/// <summary>
/// This interface defines a complete generic visitor for a parse tree produced
/// by <see cref="calcExprParser"/>.
/// </summary>
/// <typeparam name="Result">The return type of the visit operation.</typeparam>
[System.CodeDom.Compiler.GeneratedCode("ANTLR", "4.5.1")]
[System.CLSCompliant(false)]
public interface IcalcExprVisitor<Result> : IParseTreeVisitor<Result> {
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.parse"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitParse([NotNull] calcExprParser.ParseContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitExpr([NotNull] calcExprParser.ExprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.or_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitOr_expr([NotNull] calcExprParser.Or_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.or_body"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitOr_body([NotNull] calcExprParser.Or_bodyContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.and_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitAnd_expr([NotNull] calcExprParser.And_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.and_body"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitAnd_body([NotNull] calcExprParser.And_bodyContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.rel_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitRel_expr([NotNull] calcExprParser.Rel_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.eq_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitEq_expr([NotNull] calcExprParser.Eq_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.add_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitAdd_expr([NotNull] calcExprParser.Add_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.add_body"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitAdd_body([NotNull] calcExprParser.Add_bodyContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.mult_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitMult_expr([NotNull] calcExprParser.Mult_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.mult_body"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitMult_body([NotNull] calcExprParser.Mult_bodyContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.unary_expr"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitUnary_expr([NotNull] calcExprParser.Unary_exprContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>func</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitFunc([NotNull] calcExprParser.FuncContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>constVar</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstVar([NotNull] calcExprParser.ConstVarContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>variable</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitVariable([NotNull] calcExprParser.VariableContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>nullValue</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitNullValue([NotNull] calcExprParser.NullValueContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>trueValue</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitTrueValue([NotNull] calcExprParser.TrueValueContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>falseValue</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitFalseValue([NotNull] calcExprParser.FalseValueContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>braceExpr</c>
	/// labeled alternative in <see cref="calcExprParser.atom"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitBraceExpr([NotNull] calcExprParser.BraceExprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.var"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitVar([NotNull] calcExprParser.VarContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.funcname"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitFuncname([NotNull] calcExprParser.FuncnameContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>multiExpr</c>
	/// labeled alternative in <see cref="calcExprParser.params"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitMultiExpr([NotNull] calcExprParser.MultiExprContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.unaryparam"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitUnaryparam([NotNull] calcExprParser.UnaryparamContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>continueParam</c>
	/// labeled alternative in <see cref="calcExprParser.briefparam"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitContinueParam([NotNull] calcExprParser.ContinueParamContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>constNumber</c>
	/// labeled alternative in <see cref="calcExprParser.const_var"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstNumber([NotNull] calcExprParser.ConstNumberContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>constString</c>
	/// labeled alternative in <see cref="calcExprParser.const_var"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstString([NotNull] calcExprParser.ConstStringContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>constDate</c>
	/// labeled alternative in <see cref="calcExprParser.const_var"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstDate([NotNull] calcExprParser.ConstDateContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>constTime</c>
	/// labeled alternative in <see cref="calcExprParser.const_var"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstTime([NotNull] calcExprParser.ConstTimeContext context);
	/// <summary>
	/// Visit a parse tree produced by the <c>constCurrency</c>
	/// labeled alternative in <see cref="calcExprParser.const_var"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConstCurrency([NotNull] calcExprParser.ConstCurrencyContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.const_string"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitConst_string([NotNull] calcExprParser.Const_stringContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.date"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitDate([NotNull] calcExprParser.DateContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.datename"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitDatename([NotNull] calcExprParser.DatenameContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.time"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitTime([NotNull] calcExprParser.TimeContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.currency"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitCurrency([NotNull] calcExprParser.CurrencyContext context);
	/// <summary>
	/// Visit a parse tree produced by <see cref="calcExprParser.currencyunit"/>.
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	Result VisitCurrencyunit([NotNull] calcExprParser.CurrencyunitContext context);
}
