/*
 * File: helpers.go
 * File Created: Thursday, 20th June 2024 12:47:35 pm
 * Last Modified: Thursday, 20th June 2024 5:45:09 pm
 * Author: Akhil Datla
 * Copyright © Akhil Datla 2024
 */

package reportgen

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

// rangeRefToCoordinates provides a function to convert range reference to a
// pair of coordinates.
func rangeRefToCoordinates(ref string) ([]int, error) {
	rng := strings.Split(strings.ReplaceAll(ref, "$", ""), ":")
	if len(rng) < 2 {
		return nil, ErrParameterInvalid
	}
	return cellRefsToCoordinates(rng[0], rng[1])
}

// cellRefsToCoordinates provides a function to convert cell range to a
// pair of coordinates.
func cellRefsToCoordinates(firstCell, lastCell string) ([]int, error) {
	coordinates := make([]int, 4)
	var err error
	coordinates[0], coordinates[1], err = excelize.CellNameToCoordinates(firstCell)
	if err != nil {
		return coordinates, err
	}
	coordinates[2], coordinates[3], err = excelize.CellNameToCoordinates(lastCell)
	return coordinates, err
}

// sortCoordinates provides a function to correct the cell range, such
// correct C1:B3 to B1:C3.
func sortCoordinates(coordinates []int) error {
	if len(coordinates) != 4 {
		return ErrCoordinates
	}
	if coordinates[2] < coordinates[0] {
		coordinates[2], coordinates[0] = coordinates[0], coordinates[2]
	}
	if coordinates[3] < coordinates[1] {
		coordinates[3], coordinates[1] = coordinates[1], coordinates[3]
	}
	return nil
}

// checkSize validates that the data size matches the provided cell range size.
func checkSize(data SupportedDataType, coordinates []int) error {
	switch v := data.(type) {
	case Scalar:
		if v.Size() != 1 {
			return ErrSizeInvalid
		}
	case Vector:
		if v.Size() != (coordinates[3] - coordinates[1] + 1) {
			return ErrSizeInvalid
		}
	case Table:
		if v.Size() != (coordinates[3]-coordinates[1]+1)*(coordinates[2]-coordinates[0]+1) {
			return ErrSizeInvalid
		}
	}
	return nil
}
