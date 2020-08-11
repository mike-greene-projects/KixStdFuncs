;;
;;=====================================================================================-----
;;
;;FUNCTION       Ceil()
;;
;;ACTION         Round float always up to integer ceiling; Ceil(1.22) = 2
;;
;;AUTHOR         Michael Greene
;;
;;CONTRIBUTORS   Glenn Barnas
;;
;;VERSION        1.0 - 2020-04-07
;;
;;HISTORY        n/a
;;
;;SYNTAX         Floor($_Float)
;;
;;PARAMETERS     $_Float - REQUIRED - Number to be rounded
;;
;;REMARKS        
;;
;;RETURNS        Integer
;;
;;DEPENDENCIES   none
;;
;;TESTED WITH    W2K8, W2K12, W2K16
;;
;;EXAMPLES       
;
Function Ceil($_Float)

  Dim $_RoundedFloat
  
  $RoundedFloat = Round($_Float)

  IIf($RoundedFloat >= $Float, $RoundedFloat, $RoundedFloat + 1)

  Exit 0

EndFunction