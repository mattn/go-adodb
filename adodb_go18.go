// +build go1.8

package adodb

import (
	"database/sql/driver"
	"errors"

	"golang.org/x/net/context"
)

// Ping implement Pinger.
func (c *AdodbConn) Ping(ctx context.Context) error {
	if c.db == nil {
		return errors.New("Connection was closed")
	}
	return nil
}

/*
func (c *AdodbConn) QueryContext(ctx context.Context, query string, args []driver.NamedValue) (driver.Rows, error) {
	list := make([]namedValue, len(args))
	for i, nv := range args {
		list[i] = namedValue(nv)
	}
	return c.query(ctx, query, list)
}

func (c *AdodbConn) ExecContext(ctx context.Context, query string, args []driver.NamedValue) (driver.Result, error) {
	list := make([]namedValue, len(args))
	for i, nv := range args {
		list[i] = namedValue(nv)
	}
	return c.exec(ctx, query, list)
}
*/

func (c *AdodbConn) PrepareContext(ctx context.Context, query string) (driver.Stmt, error) {
	return c.prepare(ctx, query)
}

func (c *AdodbConn) BeginTx(ctx context.Context, opts driver.TxOptions) (driver.Tx, error) {
	return c.begin(ctx)
}

func (s *AdodbStmt) QueryContext(ctx context.Context, args []driver.NamedValue) (driver.Rows, error) {
	list := make([]namedValue, len(args))
	for i, nv := range args {
		list[i] = namedValue(nv)
	}
	return s.query(ctx, list)
}

func (s *AdodbStmt) ExecContext(ctx context.Context, args []driver.NamedValue) (driver.Result, error) {
	list := make([]namedValue, len(args))
	for i, nv := range args {
		list[i] = namedValue(nv)
	}
	return s.exec(ctx, list)
}
