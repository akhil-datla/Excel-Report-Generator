/*
 * File: errors.go
 * File Created: Thursday, 20th June 2024 1:15:32 pm
 * Last Modified: Thursday, 20th June 2024 5:45:06 pm
 * Author: Akhil Datla
 * Copyright © Akhil Datla 2024
 */

package reportgen

import "errors"

var (
	ErrCoordinates      = errors.New("coordinates length must be 4")
	ErrParameterInvalid = errors.New("parameter is invalid")
	ErrSizeInvalid      = errors.New("size of the data type does not match the size of the cell range")
	ErrDataTypeInvalid  = errors.New("data type is not supported")
)
