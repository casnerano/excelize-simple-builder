package esb

import "github.com/xuri/excelize/v2"

type Column struct {
	Title string
	Value func(any) any

	NestedCols Columns
}

func (c Column) CountLeaves() int {
	if len(c.NestedCols) == 0 {
		return 1
	}

	total := 0
	for _, nestedCol := range c.NestedCols {
		total += nestedCol.CountLeaves()
	}

	return total
}

type Columns []Column

func (c Columns) Depth() int {
	maxDepth := 0
	for _, col := range c {
		if depth := col.NestedCols.Depth() + 1; depth > maxDepth {
			maxDepth = depth
		}
	}

	return maxDepth
}

func Col[T, V any](header string, getter func(T) V) Column {
	return Column{
		Title: header,
		Value: func(row any) any {
			if row == nil {
				return nil
			}
			return getter(row.(T))
		},
	}
}

func Group[T, V any](header string, getter func(T) V, nestedCols ...Column) Column {
	return Column{
		Title: header,
		Value: func(row any) any {
			if row == nil {
				return nil
			}
			return getter(row.(T))
		},
		NestedCols: nestedCols,
	}
}

type StyleFunc func(style *excelize.Style)

type Styles struct {
	headerStyle StyleFunc
	bodyStyle   StyleFunc
}

type Option[T any] func(*Report[T])

func WithHeaderStyle[T any](fn StyleFunc) Option[T] {
	return func(r *Report[T]) { r.styles.headerStyle = fn }
}

func WithBodyStyle[T any](fn StyleFunc) Option[T] {
	return func(r *Report[T]) { r.styles.bodyStyle = fn }
}

type Report[T any] struct {
	columns []Column
	styles  Styles
}

func New[T any](columns []Column, opts ...Option[T]) *Report[T] {
	r := &Report[T]{columns: columns}

	for _, opt := range opts {
		opt(r)
	}

	return r
}
