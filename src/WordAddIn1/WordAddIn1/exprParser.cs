using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordAddIn1
{

// EXP = EXP [OP EXP]
// OP = '+ - * / > >= == <= <'
// EXP = PARA | FUNC( (PARA)* (,PARA)*)
// PARA = VAR | CONST_VAR
// VAR= STRING_VAR
// CONST_VAR = NUMBER_CONST | "STRING_VAR"
// STRING_VAR=[a-zA-Z]+(a-zA-Z0-9_-)*
// NUMBER_CONST= ["+"|"-"]  ( [1-9]+(0-9)*   |   [0-9]+(0-9)*[.[0-9]+(0-9)* ]  )

// TOKEN : /* IDENTIFIERS */
// {
//   < IDENTIFIER: <LETTER> (<LETTER>|<DIGIT>)* >
// |
//   < #LETTER: [ "a"-"z", "A"-"Z" ] >
// |
//   < #DIGIT: [ "0"-"9"] >
// }


    class exprParser
    {
        
    }
}
