using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using Antlr4.Runtime.Tree;
using Word = Microsoft.Office.Interop.Word;


//@TODO, 2016-01-18
// 1. support EXCEL functions
// 2. support dynamic functions loading such as DLL or lib?
// 3. VAR must exist(defined) but could be EMPTY(no value)
// 4. if var defined but empty, function could ignore it(as 0 or "", or jump over it). Like Excel.
// 5. 
// 6. 


namespace Antlr4.Parser
{
    public class EvalResult
    {
        public String Text;
        public object Value1;
        public object Value2;
        public ArrayList Values = new ArrayList();
        public String Op; // 
        public Boolean bInvalid = false;
        public String strExceptionMsg = "";

        public void clone(EvalResult other)
        {
            if (other == null)
                return;

            this.Text = other.Text;
            this.Value1 = other.Value1;
            this.Value2 = other.Value2;
            this.Op = other.Op;
            this.bInvalid = other.bInvalid;
            this.strExceptionMsg = other.strExceptionMsg;

            if (other.Values != null && other.Values.Count > 0)
            {
                for (int i = 0; i < other.Values.Count; i++)
                    this.Values.Add(other.Values[i]);
            }
        }
    }

    public class EvalExprVisitor : calcExprBaseVisitor<EvalResult>
    {
        protected Hashtable m_hashAllVars = null;
        public Hashtable m_hashRefVars = null;

        public void SetRefVarsHash(Hashtable oRefVarsHash)
        {
            m_hashRefVars = oRefVarsHash;
        }

        public void SetVarHash(Hashtable oHashTbl)
        {
            m_hashAllVars = oHashTbl;
        }

        public override EvalResult VisitParse(calcExprParser.ParseContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            m_hashRefVars.Clear();

            calcExprParser.ExprContext exprCxt = context.expr();

            if (exprCxt == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitExpr(exprCxt);//VisitChildren(context);

            if (retSub == null)
            {
                return null;
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitExpr(calcExprParser.ExprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            calcExprParser.Or_exprContext orCxt = context.or_expr();

            if(orCxt == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitOr_expr(orCxt);//VisitChildren(context);

            if (retSub == null)
            {
                return null;
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitOr_expr(calcExprParser.Or_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            retSub = VisitAnd_expr(context.left);
            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            calcExprParser.Or_bodyContext[] ctxs = context.or_body();
            for (int i = 0; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                retSub1 = VisitOr_body(ctxs[i]);

                if (retSub1 == null)
                    continue;

                if (retSub1.Op != null)
                {
                    if (retSub1.bInvalid)
                    {
                        ret.clone(retSub1);
                        return ret;
                    }

                    if (retSub1.Op.Equals("OR"))
                    {
                        retSub.Text += " OR " + retSub1.Text;

                        Boolean bRet = false, bRet1 = false;

                        try
                        {
                            bRet = System.Convert.ToBoolean(retSub.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非Boolean值:" + retSub.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }

                        try
                        {
                            bRet1 = System.Convert.ToBoolean(retSub1.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非Boolean值:" + retSub1.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }

                        try
                        {
                            retSub.Value1 = (bRet || bRet1);

                            // 
                            if (bRet || bRet1)
                            {
                                break;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非Boolean值:" + retSub.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                }
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitOr_body(calcExprParser.Or_bodyContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitAnd_expr(context.and_expr());

            if (context.op.Type == calcExprParser.OR)
            {
                retSub.Op = "OR";
            }

            if (retSub == null)
            {
                return null;
            }

            ret.clone(retSub);
            return ret;
        }


        public override EvalResult VisitAnd_expr(calcExprParser.And_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            retSub = VisitRel_expr(context.left);

            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            calcExprParser.And_bodyContext[] ctxs = context.and_body();
            for (int i = 0; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                retSub1 = VisitAnd_body(ctxs[i]);

                if (retSub1 == null)
                    continue;

                if (retSub1.Op != null)
                {
                    if (retSub1.bInvalid)
                    {
                        ret.clone(retSub1);
                        return ret;
                    }

                    if (retSub1.Op.Equals("AND"))
                    {
                        retSub.Text += " AND " + retSub1.Text;

                        Boolean bRet = false, bRet1 = false;

                        try
                        {
                            bRet = System.Convert.ToBoolean(retSub.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非Boolean值:" + retSub.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }

                        try
                        {
                            bRet1 = System.Convert.ToBoolean(retSub1.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非Boolean值:" + retSub1.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }

                        try
                        {
                            retSub.Value1 = (bRet && bRet1);

                            if (!(bRet && bRet1))
                            {
                                break;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非Boolean值:" + retSub.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                }
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitAnd_body(calcExprParser.And_bodyContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitRel_expr(context.rel_expr());

            if (context.op.Type == calcExprParser.AND)
            {
                retSub.Op = "AND";
            }

            ret.clone(retSub);
            return ret;
        }


        public override EvalResult VisitRel_expr(calcExprParser.Rel_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            retSub = VisitEq_expr(context.left);
            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            double dbVal = 0.0, dbVal1 = 0.0;

            calcExprParser.Eq_exprContext[] ctxs = context.eq_expr();
            for (int i = 1; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                dbVal = dbVal1 = 0.0;

                retSub1 = VisitEq_expr(ctxs[i]);

                if (retSub1 == null)
                    continue;

                if (retSub1.bInvalid)
                {
                    ret.clone(retSub1);
                    return ret;
                }

                if (context.op != null)
                {
                    try
                    {
                        dbVal = System.Convert.ToDouble(retSub.Value1);
                    }
                    catch (System.Exception ex)
                    {
                        retSub.strExceptionMsg = "#INVALID:非数字：" + retSub.Value1;
                        retSub.bInvalid = true;
                        ret.clone(retSub);
                        return ret;
                    }
                    finally
                    {
                    }


                    try
                    {
                        dbVal1 = System.Convert.ToDouble(retSub1.Value1);
                    }
                    catch (System.Exception ex)
                    {
                        retSub.strExceptionMsg = "#INVALID:非数字：" + retSub1.Value1;
                        retSub.bInvalid = true;
                        ret.clone(retSub);
                        return ret;
                    }
                    finally
                    {
                    }

                    if (context.op.Type == calcExprParser.LT)
                    {
                        retSub.Text = retSub.Text + " < " + retSub1.Text;

                        try
                        {
                            retSub.Value1 = (dbVal < dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                    else if (context.op.Type == calcExprParser.GT)
                    {
                        retSub.Text = retSub.Text + " > " + retSub1.Text;

                        try
                        {
                            retSub.Value1 = (dbVal > dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                    else if (context.op.Type == calcExprParser.LEQ)
                    {
                        retSub.Text = retSub.Text + " <= " + retSub1.Text;
                        try
                        {
                            retSub.Value1 = (dbVal <= dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                    else if (context.op.Type == calcExprParser.GEQ)
                    {
                        retSub.Text = retSub.Text + " >= " + retSub1.Text;
                        try
                        {
                            retSub.Value1 = (dbVal >= dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                }
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitEq_expr(calcExprParser.Eq_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            retSub = VisitAdd_expr(context.left);
            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            calcExprParser.Add_exprContext[] ctxs = context.add_expr();
            for (int i = 1; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                retSub1 = VisitAdd_expr(ctxs[i]);

                if (retSub1 == null)
                    continue;

                if (retSub1.bInvalid)
                {
                    ret.clone(retSub1);
                    return ret;
                }

                if (context.op != null)
                {
                    if (context.op.Type == calcExprParser.EQ)
                    {
                        retSub.Text = retSub.Text + " == " + retSub1.Text;
                        retSub.Value1 = (retSub.Value1 == retSub1.Value1);
                    }
                    else if (context.op.Type == calcExprParser.NEQ)
                    {
                        retSub.Text = retSub.Text + " != " + retSub1.Text;
                        retSub.Value1 = (retSub.Value1 != retSub1.Value1);
                    }
                }

            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitAdd_expr(calcExprParser.Add_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            retSub = VisitMult_expr(context.left);

            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            calcExprParser.Add_bodyContext[] ctxs = context.add_body();
            double dbVal = 0.0, dbVal1 = 0.0;

            for (int i = 0; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                retSub1 = VisitAdd_body(ctxs[i]);
                dbVal = dbVal1 = 0.0;

                if (retSub1 == null)
                    continue;

                if (retSub1.Op != null)
                {
                    if (retSub1.bInvalid)
                    {
                        ret.clone(retSub1);
                        return ret;
                    }

                    try
                    {
                        dbVal = System.Convert.ToDouble(retSub.Value1);
                    }
                    catch (System.Exception ex)
                    {
                        retSub.strExceptionMsg = "#INVALID:非数字：" + retSub.Value1;
                        retSub.bInvalid = true;

                        ret.clone(retSub);
                        return ret;
                    }
                    finally
                    {
                    }


                    try
                    {
                        dbVal1 = System.Convert.ToDouble(retSub1.Value1);
                    }
                    catch (System.Exception ex)
                    {
                        retSub.strExceptionMsg = "#INVALID:非数字：" + retSub1.Value1;
                        retSub.bInvalid = true;
                        ret.clone(retSub);
                        return ret;
                    }
                    finally
                    {
                    }

                    if (retSub1.Op.Equals("PLUS"))
                    {
                        retSub.Text = retSub.Text + " + " + retSub1.Text;

                        try
                        {
                            retSub.Value1 = (dbVal + dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }

                    }
                    else if (retSub1.Op.Equals("MINUS"))
                    {
                        retSub.Text = retSub.Text + " - " + retSub1.Text;

                        try
                        {
                            retSub.Value1 = (dbVal - dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }

                    }
                }

            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitAdd_body(calcExprParser.Add_bodyContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitMult_expr(context.mult_expr());

            if (retSub == null)
            {
                return null;
            }

            if (context.op.Type == calcExprParser.PLUS)
            {
                retSub.Op = "PLUS";
            }
            else if (context.op.Type == calcExprParser.MINUS)
            {
                retSub.Op = "MINUS";
            }

            ret.clone(retSub);
            return ret;
        }


        public override EvalResult VisitMult_expr(calcExprParser.Mult_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            retSub = VisitUnary_expr(context.left);

            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            double dbVal = 0.0, dbVal1 = 0.0;
            calcExprParser.Mult_bodyContext[] ctxs = context.mult_body();

            for (int i = 0; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                retSub1 = VisitMult_body(ctxs[i]);

                if (retSub1 == null)
                    continue;

                if (retSub1.Op != null)
                {
                    if (retSub1.bInvalid)
                    {
                        ret.clone(retSub1);
                        return ret;
                    }

                    try
                    {
                        dbVal = System.Convert.ToDouble(retSub.Value1);
                    }
                    catch (System.Exception ex)
                    {
                        retSub.strExceptionMsg = "#INVALID:非数字：" + retSub.Value1;
                        retSub.bInvalid = true;

                        ret.clone(retSub);
                        return ret;
                    }
                    finally
                    {
                    }


                    if (retSub1.Op.Equals("MULT"))
                    {
                        retSub.Text = retSub.Text + " * " + retSub1.Text;
                        try
                        {
                            retSub.Value1 = (dbVal * System.Convert.ToDouble(retSub1.Value1));
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非数字：" + retSub1.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                    }
                    else if (retSub1.Op.Equals("DIV"))
                    {
                        retSub.Text = retSub.Text + " / " + retSub1.Text;

                        try
                        {
                            dbVal1 = System.Convert.ToDouble(retSub1.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非数字：" + retSub1.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }


                        if (dbVal1 == (double)0.0)
                        {
                            retSub.strExceptionMsg = "#INVALID:被0除";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }


                        try
                        {
                            retSub.Value1 = (dbVal / dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                    else if (retSub1.Op.Equals("MOD"))
                    {
                        retSub.Text = retSub.Text + " % " + retSub1.Text;

                        try
                        {
                            dbVal1 = System.Convert.ToDouble(retSub1.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID:非数字：" + retSub1.Value1;
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }


                        if (dbVal1 == (double)0.0)
                        {
                            retSub.strExceptionMsg = "#INVALID:被0除";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }

                        try
                        {
                            retSub.Value1 = (dbVal % dbVal1);
                        }
                        catch (System.Exception ex)
                        {
                            retSub.strExceptionMsg = "#INVALID";
                            retSub.bInvalid = true;
                            ret.clone(retSub);
                            return ret;
                        }
                        finally
                        {
                        }
                    }
                }

            }

            ret.clone(retSub);
            return ret;
        }


        public override EvalResult VisitMult_body(calcExprParser.Mult_bodyContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitUnary_expr(context.unary_expr());

            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            if (context.op.Type == calcExprParser.MULT)
            {
                retSub.Op = "MULT";
            }
            else if (context.op.Type == calcExprParser.DIV)
            {
                retSub.Op = "DIV";
            }
            else if (context.op.Type == calcExprParser.MOD)
            {
                retSub.Op = "MOD";
            }

            ret.clone(retSub);
            return ret;
        }


        public override EvalResult VisitUnary_expr(calcExprParser.Unary_exprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = Visit(context.atom());

            if (retSub == null)
            {
                return null;
            }
                
            if(retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            if (context.op != null && context.op.Type == calcExprParser.MINUS)
            {
                try
                {
                    retSub.Value1 = -1 * System.Convert.ToDouble(retSub.Value1); // ?? 
                }
                catch (System.Exception ex)
                {
                    retSub.strExceptionMsg = "#INVALID:非数字：" + retSub.Value1;
                    retSub.bInvalid = true;
                }
                finally
                {
                }
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitFunc(calcExprParser.FuncContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retParams = null;

            // EvalResult retFuncName = new EvalResult();

            String strFuncName = context.funcname.Text.ToUpper();
            ret.Text = strFuncName;
            ret.Value1 = strFuncName;

            // retFuncName.Text = context.funcname.Text;
            // retFuncName.Value1 = retFuncName.Text;

            calcExprParser.ParamsContext paramsCtx = context.@params();

            if (paramsCtx == null) // no params
            {
                // ret.clone(retFuncName);
            }
            else
            {
                retParams = Visit(paramsCtx);

                if (retParams == null)
                {
                    return null;
                }
                    
                if(retParams.bInvalid)
                {
                    ret.clone(retParams);
                    return ret;
                }

                ret.Text = strFuncName + retParams.Text;
                ret.Value1 = strFuncName;
            }

            EvalResult objRet = null;

            if (strFuncName.Equals("SUM"))
            {
                objRet = FuncSum(retParams);
                objRet.Text = context.GetText();
            }
            else if (strFuncName.Equals("AVG") || strFuncName.Equals("AVERAGE"))
            {
                objRet = FuncAvg(retParams);
                objRet.Text = context.GetText();
            }
            else if (strFuncName.Equals("CONCAT") || strFuncName.Equals("CONCATENATE"))
            {
                objRet = FuncConcateNate(retParams);
                objRet.Text = context.GetText();
            }
            else if (strFuncName.Equals("CNT") || strFuncName.Equals("COUNT"))
            {
                objRet = FuncCount(retParams);
                objRet.Text = context.GetText();
            }
            else if (strFuncName.Equals("CNTA") || strFuncName.Equals("COUNTA"))
            {
                objRet = FuncCountA(retParams);
                objRet.Text = context.GetText();
            }
            else if (strFuncName.Equals("FMT") || strFuncName.Equals("FORMAT"))
            {
                objRet = FuncFormat(retParams);
                objRet.Text = context.GetText();
            }
            else
            {
                objRet = new EvalResult();
                objRet.Text = context.GetText();
                objRet.strExceptionMsg = "#INVALID:未支持函数:" + ret.Text;
                objRet.bInvalid = true;
            }

            ret.clone(objRet);
            return ret;
        }

        //         public override EvalResult VisitFuncname(calcExprParser.FuncnameContext context)
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }

        public override EvalResult VisitMultiExpr(calcExprParser.MultiExprContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null, retSub1 = null;

            calcExprParser.UnaryparamContext[] ctxs = context.unaryparam();
            for (int i = 0; ctxs != null && i < ctxs.GetLength(0); i++)
            {
                retSub = VisitUnaryparam(ctxs[i]);

                if (retSub == null)
                    continue;

                if (retSub.bInvalid)
                {
                    ret.clone(retSub);
                    return ret;
                }

                retSub1 = new EvalResult();
                retSub1.clone(retSub);

                ret.Values.Add(retSub1);
            }

            return ret;
        }

        public override EvalResult VisitUnaryparam(calcExprParser.UnaryparamContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = null;

            calcExprParser.ExprContext exprCtx = context.expr();
            calcExprParser.BriefparamContext briefParamCtx = context.briefparam();

            if (exprCtx != null)
            {
                retSub = VisitExpr(exprCtx);
            }
            else if (briefParamCtx != null)
            {
                retSub = Visit(briefParamCtx);
            }

            if (retSub == null)
            {
                return null;
            }
               
            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitContinueParam(calcExprParser.ContinueParamContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = Visit(context.var());

            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            int nStart = 0, nEnd = 0;

            try
            {
                nStart = System.Convert.ToInt16(context.sInx.Text);
                retSub.Value1 = nStart;
            }
            catch (System.Exception ex)
            {
                retSub.strExceptionMsg = "#INVALID:非整数:" + context.sInx.Text;
                retSub.bInvalid = true;
                return retSub;
            }

            try
            {
                nEnd = System.Convert.ToInt16(context.eInx.Text);
                retSub.Value2 = nEnd;
            }
            catch (System.Exception ex)
            {
                retSub.strExceptionMsg = "#INVALID:非整数:" + context.eInx.Text;
                retSub.bInvalid = true;

                ret.clone(retSub);
                return ret;
            }

            for (int i = nStart; i <= nEnd; i++)
            {
                m_hashRefVars[retSub.Text + i] = (retSub.Text + i);
            }

            ret.clone(retSub);

            return ret;
        }

        public override EvalResult VisitVariable(calcExprParser.VariableContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            EvalResult retSub = VisitVar(context.var());

            if (retSub == null)
            {
                return null;
            }

            if (retSub.bInvalid)
            {
                ret.clone(retSub);
                return ret;
            }

            // keep unique
            m_hashRefVars[retSub.Text] = retSub.Value1.ToString();

            if (m_hashAllVars != null)
            {
                Word.ContentControl defCtrl = (Word.ContentControl)m_hashAllVars[retSub.Text];
                if (defCtrl != null)
                {
                    retSub.Value1 = defCtrl.Range.Text;
                }
                else
                {
                    retSub.strExceptionMsg = "#INVALID:变量未定义:" + retSub.Text;
                    // ret.strExceptionMsg = "#INVALID:变量未定义:" + ret.Text;
                    // ret.bInvalid = true;
                    // return ret1;
                }
            }

            ret.clone(retSub);
            return ret;
        }

        public override EvalResult VisitNullValue(calcExprParser.NullValueContext context)
        {
            EvalResult ret = new EvalResult();
            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ret.Text = context.GetText();
            ret.Value1 = ret.Text; 

            return ret;
        }

        public override EvalResult VisitTrueValue(calcExprParser.TrueValueContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ret.Text = context.GetText();
            ret.Value1 = true;

            return ret;
        }

        public override EvalResult VisitFalseValue(calcExprParser.FalseValueContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ret.Text = context.GetText();
            ret.Value1 = false;

            return ret;
        }


        public override EvalResult VisitBraceExpr(calcExprParser.BraceExprContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            calcExprParser.ExprContext exprCxt = context.expr();

            if(exprCxt == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            // should calc at first
            EvalResult ret1 = VisitExpr(exprCxt);// VisitChildren(context);

            if (ret1 != null)
            {
                ret.clone(ret1);
            }
            else
            {
                return null;
            }

            return ret;
        }


        public override EvalResult VisitVar(calcExprParser.VarContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ITerminalNode varNode = context.VARID();

            if (varNode == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ret.Text = varNode.GetText();
            ret.Value1 = ret.Text;

            return ret;
        }

        public override EvalResult VisitConstNumber(calcExprParser.ConstNumberContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ITerminalNode numNode = context.NUMBER();
            if (numNode == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

                
            ret.Text = numNode.GetText();
            try
            {
                ret.Value1 = System.Convert.ToDouble(ret.Text);
            }
            catch (System.Exception ex)
            {
                ret.strExceptionMsg = "#INVALID:非数字:" + ret.Text;
                ret.bInvalid = true;
            }
            finally
            {
            }

            return ret;
        }

        public override EvalResult VisitConstString(calcExprParser.ConstStringContext context)
        {
            EvalResult ret = new EvalResult();

            if (context == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            String strCnt = context.GetText();

            if (strCnt == null)
            {
                ret.bInvalid = true;
                ret.strExceptionMsg = "#INVALID:解析异常，请检查输入";
                return ret;
            }

            ret.Text = strCnt;//strCnt.Replace("\"", "");
            ret.Value1 = ret.Text;

            return ret;
        }

        //         public override EvalResult VisitConstDate(calcExprParser.ConstDateContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text  = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }


        //         public override EvalResult VisitConstTime(calcExprParser.ConstTimeContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text  = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }


        //         public override EvalResult VisitConstCurrency(calcExprParser.ConstCurrencyContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text  = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret; 
        //         }

        //         public override EvalResult VisitConst_string(calcExprParser.Const_stringContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text = context.STRING().GetText().Replace("\"", "");
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }

        //         public override EvalResult VisitDate(calcExprParser.DateContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text  = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }


        //         public override EvalResult VisitDatename(calcExprParser.DatenameContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text  = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }

        //         public override EvalResult VisitTime(calcExprParser.TimeContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text    = context.GetText();
        //             ret.Value1   = ret.Text;
        // 
        //             return ret;
        //         }

        //         public override EvalResult VisitCurrency(calcExprParser.CurrencyContext context)
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }


        //         public override EvalResult VisitCurrencyunit(calcExprParser.CurrencyunitContext context) 
        //         {
        //             EvalResult ret = new EvalResult();
        // 
        //             ret.Text = context.GetText();
        //             ret.Value1 = ret.Text;
        // 
        //             return ret;
        //         }

        public EvalResult FuncSum(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            double ret = 0.0;
            int nStart = 0, nEnd = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    if (!defCtrl.Range.Text.Trim().Equals(""))
                                    {
                                        ret += Convert.ToDouble(defCtrl.Range.Text.Trim());
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID:变量(" + (newEv.Text + i) + ")值非数字:" + defCtrl.Range.Text;
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }// for
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    String strValue = newEv.Value1.ToString();
                    if (!strValue.Trim().Equals(""))
                    {
                        try
                        {
                            ret += Convert.ToDouble(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非数字:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    }

                }// else
            }

            objRet.Value1 = ret;
            return objRet;
        }


        public EvalResult FuncAvg(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            double ret = 0.0;
            int nStart = 0, nEnd = 0, nTotal = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }


                        // nTotal += nEnd - nStart + 1;

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    if (!defCtrl.Range.Text.Trim().Equals(""))
                                    {
                                        ret += Convert.ToDouble(defCtrl.Range.Text.Trim());
                                        nTotal++;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID:变量(" + (newEv.Text + i) + ")非数字:" + defCtrl.Range.Text;
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    String strValue = newEv.Value1.ToString();
                    if (!strValue.Trim().Equals(""))
                    {
                        try
                        {
                            ret += Convert.ToDouble(newEv.Value1);
                            nTotal++;
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非数字:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    }

                }
            }

            if (nTotal == 0)
            {
                objRet.strExceptionMsg = "#INVALID:被0除";
                objRet.bInvalid = true;
                return objRet;
            }

            ret = ret / nTotal;

            objRet.Value1 = ret;
            return objRet;
        }


        public EvalResult FuncConcateNate(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            String ret = "";
            int nStart = 0, nEnd = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    ret += defCtrl.Range.Text;
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID";
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    //String strValue = newEv.Value1.ToString();
                    //if (!strValue.Trim().Equals(""))
                    //{
                        try
                        {
                            ret += newEv.Value1;
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非法:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    //}
                }
            }

            objRet.Value1 = ret;
            return objRet;
        }


        public EvalResult FuncCount(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            uint ret = 0;
            int nStart = 0, nEnd = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    String strVal = defCtrl.Range.Text.Trim();
                                    double dbVal = 0.0;

                                    if (!strVal.Equals("") && double.TryParse(strVal, out dbVal))
                                    {
                                        ret += 1;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID";
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    String strValue = newEv.Value1.ToString();
                    double dbTmp = 0.0;
                    if (!strValue.Trim().Equals(""))
                    {
                        try
                        {
                            if (double.TryParse(strValue, out dbTmp))
                            {
                                ret += 1;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非数字:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    }

                }
            }

            objRet.Value1 = ret;
            return objRet;
        }


        public EvalResult FuncCountA(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            uint ret = 0;
            int nStart = 0, nEnd = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    String strVal = defCtrl.Range.Text.Trim();
                                    if (!strVal.Equals(""))
                                    {
                                        ret += 1;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID";
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    String strValue = newEv.Value1.ToString();
                    if (!strValue.Trim().Equals(""))
                    {
                        try
                        {
                            ret += 1;
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非数字:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    }

                }
            }

            objRet.Value1 = ret;
            return objRet;
        }


        // abs,cos,sin,
        // max,min
        // format()
        // 

        // 1:value
        // 2:format string, like:###.##,0.00
        public EvalResult FuncFormat(EvalResult param)
        {
            EvalResult objRet = new EvalResult();

            if (param.Values.Count < 2)
            {
                objRet.strExceptionMsg = "#INVALID:参数太少";
                objRet.bInvalid = true;

                return objRet;
            }

            EvalResult valEv = (EvalResult)param.Values[0];
            EvalResult fmtEv = (EvalResult)param.Values[1];

            String strFmt0 = (String)fmtEv.Value1, strFmt = "";

            strFmt = "{0:" + strFmt0.Replace("\"","") + "}"; // "{0:0.##}"

            String strValueFmted = "";

            double dbVal = 0.0;


            try
            {
                dbVal = Convert.ToDouble(valEv.Value1);
                try
                {
                    strValueFmted = String.Format(strFmt, dbVal);

                    objRet.Text = strValueFmted;
                    objRet.Value1 = strValueFmted;
                }
                catch (System.Exception ex)
                {
                    objRet.strExceptionMsg = "#INVALID:格式化参数错误：" + strFmt0;
                    objRet.bInvalid = true;

                    return objRet;
                }
                finally
                {

                }
            }
            catch (System.Exception ex)
            {
                objRet.strExceptionMsg = "#INVALID:非数字：" + valEv.Value1;
                objRet.bInvalid = true;

                return objRet;
            }
            finally
            {

            }

            return objRet;

        }

        // 
        public EvalResult FuncMax(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            double ret = 0.0, dbTmp = 0.0;
            int nStart = 0, nEnd = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    if (!defCtrl.Range.Text.Trim().Equals(""))
                                    {
                                        dbTmp = Convert.ToDouble(defCtrl.Range.Text.Trim());

                                        if (dbTmp > ret)
                                        {
                                            ret = dbTmp;
                                        }

                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID:变量(" + (newEv.Text + i) + ")值非数字:" + defCtrl.Range.Text;
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }// for
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    String strValue = newEv.Value1.ToString();
                    if (!strValue.Trim().Equals(""))
                    {
                        try
                        {
                            dbTmp = Convert.ToDouble(newEv.Value1);

                            if (dbTmp > ret)
                            {
                                ret = dbTmp;
                            }

                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非数字:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    }

                }// else
            }

            objRet.Value1 = ret;
            return objRet;
        }

        // 
        public EvalResult FuncMin(EvalResult param)
        {
            EvalResult objRet = new EvalResult();
            double ret = 0.0, dbTmp = 0.0;
            int nStart = 0, nEnd = 0;

            for (int nCnt = 0; nCnt < param.Values.Count; nCnt++)
            {
                EvalResult newEv = (EvalResult)param.Values[nCnt];

                if (newEv.bInvalid)
                {
                    return newEv;
                }

                nStart = nEnd = 0;

                if (newEv.Value2 != null)
                {
                    if (m_hashAllVars != null)
                    {
                        Word.ContentControl defCtrl = null;

                        try
                        {
                            nStart = Convert.ToInt16(newEv.Value1);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value1;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        // 
                        try
                        {
                            nEnd = Convert.ToInt16(newEv.Value2);
                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:非整数:" + newEv.Value2;
                            objRet.bInvalid = true;
                            return objRet;
                        }
                        finally
                        {
                        }

                        for (int i = nStart; i <= nEnd; i++)
                        {
                            defCtrl = (Word.ContentControl)m_hashAllVars[newEv.Text + i];
                            if (defCtrl != null)
                            {
                                try
                                {
                                    if (!defCtrl.Range.Text.Trim().Equals(""))
                                    {
                                        dbTmp = Convert.ToDouble(defCtrl.Range.Text.Trim());

                                        if (dbTmp < ret)
                                        {
                                            ret = dbTmp;
                                        }

                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    objRet.strExceptionMsg = "#INVALID:变量(" + (newEv.Text + i) + ")值非数字:" + defCtrl.Range.Text;
                                    objRet.bInvalid = true;
                                    return objRet;
                                }
                            }
                            else
                            {
                                continue;
                                objRet.strExceptionMsg = "#INVALID:变量未定义:" + (newEv.Text + i);
                                objRet.bInvalid = true;
                                return objRet;
                            }
                        }// for
                    }
                    else
                    {
                        objRet.strExceptionMsg = "#NOVAR";
                        objRet.bInvalid = true;
                        return objRet;
                    }
                }
                else
                {
                    String strValue = newEv.Value1.ToString();
                    if (!strValue.Trim().Equals(""))
                    {
                        try
                        {
                            dbTmp = Convert.ToDouble(newEv.Value1);

                            if (dbTmp < ret)
                            {
                                ret = dbTmp;
                            }

                        }
                        catch (System.Exception ex)
                        {
                            objRet.strExceptionMsg = "#INVALID:变量(" + newEv.Text + ")值非数字:" + newEv.Value1;
                            if (!newEv.strExceptionMsg.Equals(""))
                            {
                                objRet.strExceptionMsg = newEv.strExceptionMsg;
                            }
                            objRet.bInvalid = true;
                            return objRet;
                        }// try
                    }

                }// else
            }

            objRet.Value1 = ret;
            return objRet;
        }



        // 
        public EvalResult FuncSin(EvalResult param)
        {

            return null;
        }

        // 
        public EvalResult FuncCos(EvalResult param)
        {

            return null;
        }


        // abs



    }


}

