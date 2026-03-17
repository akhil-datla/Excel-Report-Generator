package xlreport

// Formula represents an Excel formula string. Use this as a cell value
// to set a formula instead of a literal value.
//
//	s.Cell("B5", xlreport.Formula("SUM(B2:B4)"))
//	s.Cell("C1", xlreport.Formula("IF(A1>0,\"Yes\",\"No\")"))
//	s.Cell("D1", xlreport.Formula("VLOOKUP(A1,Data!A:B,2,FALSE)"))
type Formula string
