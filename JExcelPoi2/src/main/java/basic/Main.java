package basic;

import kotlins.*;

public class Main {

    public static void main(String[] args) {
        Data.inputData();

        // JAVA
        NegativeCell.checkNegative();
        DecimalCell.checkDecimal();
        ErrCells.checkErrCells();
        ErrorCell.checkErrCell();
        StringCell.checkString();
        NullCell.checkNull();
        ZeroCell.checkZero();

        // KOTLIN
//        ZeroCells.Companion.checkZero();
//        ErrCell.Companion.checkErrCell();
//        ErrorCells.Companion.checkErrCells();
//        StrCell.Companion.checkString();
//        NegativeCells.Companion.checkNegative();
//        NullCells.Companion.checkNull();
//        DecCells.Companion.checkDecimal();
    }
}
