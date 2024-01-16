package basic;

import kotlins.*;

public class Main {

    public static void main(String[] args) {
        Data.inputData();

        // JAVA
        ZeroCell.checkZero();
        ErrCells.checkErrCells();
        ErrorCell.checkErrCell();
        StringCell.checkString();
        NegativeCell.checkNegative();
        NullCell.checkNull();
        DecimalCell.checkDecimal();

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
