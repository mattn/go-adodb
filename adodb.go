package adodb

import (
	"database/sql"
	"database/sql/driver"
	"io"
	"math"
	"math/big"
	"reflect"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/net/context"
)

func init() {
	sql.Register("adodb", &AdodbDriver{})
}

type AdodbDriver struct {
}

type AdodbConn struct {
	db *ole.IDispatch
}

type AdodbTx struct {
	c *AdodbConn
}

func (tx *AdodbTx) Commit() error {
	rv, err := oleutil.CallMethod(tx.c.db, "CommitTrans")
	if err != nil {
		return err
	}
	rv.Clear()
	return nil
}

func (tx *AdodbTx) Rollback() error {
	rv, err := oleutil.CallMethod(tx.c.db, "Rollback")
	if err != nil {
		return err
	}
	rv.Clear()
	return nil
}

func (c *AdodbConn) exec(ctx context.Context, query string, args []namedValue) (driver.Result, error) {
	s, err := c.Prepare(query)
	if err != nil {
		return nil, err
	}
	result, err := s.(*AdodbStmt).exec(ctx, args)
	s.Close()
	if err != nil && err != driver.ErrSkip {
		return nil, err
	}
	return result, nil
}

/*
func (c *AdodbConn) Query(query string, args []driver.Value) (driver.Rows, error) {
	list := make([]namedValue, len(args))
	for i, v := range args {
		list[i] = namedValue{
			Ordinal: i + 1,
			Value:   v,
		}
	}
	return c.query(context.Background(), query, list)
}

func (c *AdodbConn) query(ctx context.Context, query string, args []namedValue) (driver.Rows, error) {
	s, err := c.Prepare(query)
	if err != nil {
		return nil, err
	}
	rows, err := s.(*AdodbStmt).query(ctx, args)
	if err != nil && err != driver.ErrSkip {
		s.Close()
		return nil, err
	}
	return rows, nil
}
*/

func (c *AdodbConn) Begin() (driver.Tx, error) {
	return c.begin(context.Background())
}

func (c *AdodbConn) begin(ctx context.Context) (driver.Tx, error) {
	rv, err := oleutil.CallMethod(c.db, "BeginTrans")
	if err != nil {
		return nil, err
	}
	rv.Clear()
	return &AdodbTx{c}, nil
}

func (d *AdodbDriver) Open(dsn string) (driver.Conn, error) {
	ole.CoInitialize(0)

	unknown, err := oleutil.CreateObject("ADODB.Connection")
	if err != nil {
		return nil, err
	}
	defer unknown.Release()
	db, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	rc, err := oleutil.CallMethod(db, "Open", dsn)
	if err != nil {
		return nil, err
	}
	rc.Clear()
	return &AdodbConn{db}, nil
}

func (c *AdodbConn) Close() error {
	rv, err := oleutil.CallMethod(c.db, "Close")
	if err != nil {
		return err
	}
	rv.Clear()
	c.db.Release()
	c.db = nil
	ole.CoUninitialize()
	return nil
}

type AdodbStmt struct {
	c  *AdodbConn
	s  *ole.IDispatch
	ps *ole.IDispatch
	b  []string
}

func (c *AdodbConn) Prepare(query string) (driver.Stmt, error) {
	return c.prepare(context.Background(), query)
}

func (c *AdodbConn) prepare(ctx context.Context, query string) (driver.Stmt, error) {
	unknown, err := oleutil.CreateObject("ADODB.Command")
	if err != nil {
		return nil, err
	}
	s, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	rv, err := oleutil.PutProperty(s, "ActiveConnection", c.db)
	if err != nil {
		return nil, err
	}
	rv.Clear()
	rv, err = oleutil.PutProperty(s, "CommandText", query)
	if err != nil {
		return nil, err
	}
	rv.Clear()
	rv, err = oleutil.PutProperty(s, "CommandType", 1)
	if err != nil {
		return nil, err
	}
	rv.Clear()
	rv, err = oleutil.PutProperty(s, "Prepared", true)
	if err != nil {
		return nil, err
	}
	rv.Clear()
	val, err := oleutil.GetProperty(s, "Parameters")
	if err != nil {
		return nil, err
	}
	defer val.Clear()
	return &AdodbStmt{c, s, val.ToIDispatch(), nil}, nil
}

func (s *AdodbStmt) Bind(bind []string) error {
	s.b = bind
	return nil
}

func (s *AdodbStmt) Close() error {
	rv, err := oleutil.PutProperty(s.s, "ActiveConnection", nil)
	if err != nil {
		return err
	}
	rv.Clear()
	s.ps.Release()
	s.ps = nil
	s.s.Release()
	s.s = nil
	s.c = nil
	return nil
}

func (s *AdodbStmt) NumInput() int {
	if s.b != nil {
		return len(s.b)
	}
	rv, err := oleutil.CallMethod(s.ps, "Refresh")
	if err != nil {
		return -1
	}
	rv.Clear()
	rv, err = oleutil.GetProperty(s.ps, "Count")
	if err != nil {
		return -1
	}
	defer rv.Clear()
	return int(rv.Val)
}

func (s *AdodbStmt) bind(args []namedValue) error {
	if s.b != nil {
		for i, v := range args {
			var b string = "?"
			if len(s.b) < i {
				b = s.b[i]
			}
			unknown, err := oleutil.CallMethod(s.s, "CreateParameter", b, 12, 1)
			if err != nil {
				return err
			}
			param := unknown.ToIDispatch()
			unknown.Clear()
			rv, err := oleutil.PutProperty(param, "Value", v.Value)
			if err != nil {
				param.Release()
				unknown.Clear()
				return err
			}
			rv.Clear()
			rv, err = oleutil.CallMethod(s.ps, "Append", param)
			if err != nil {
				param.Release()
				unknown.Clear()
				return err
			}
			rv.Clear()
			param.Release()
			param.Release()
		}
	} else {
		for i, v := range args {
			var err error
			var val *ole.VARIANT
			if v.Name != "" {
				val, err = oleutil.CallMethod(s.ps, "Item", v.Name)
			} else {
				val, err = oleutil.CallMethod(s.ps, "Item", int32(i))
			}
			if err != nil {
				return err
			}
			item := val.ToIDispatch()
			val.Clear()
			rv, err := oleutil.PutProperty(item, "Value", v.Value)
			if err != nil {
				item.Release()
				return err
			}
			rv.Clear()
			item.Release()
		}
	}
	return nil
}

type namedValue struct {
	Name    string
	Ordinal int
	Value   driver.Value
}

func (s *AdodbStmt) Query(args []driver.Value) (driver.Rows, error) {
	list := make([]namedValue, len(args))
	for i, v := range args {
		list[i] = namedValue{
			Ordinal: i + 1,
			Value:   v,
		}
	}
	return s.query(context.Background(), list)
}

func (s *AdodbStmt) query(ctx context.Context, args []namedValue) (driver.Rows, error) {
	if err := s.bind(args); err != nil {
		return nil, err
	}
	rc, err := oleutil.CallMethod(s.s, "Execute")
	if err != nil {
		return nil, err
	}
	return &AdodbRows{s, rc.ToIDispatch(), -1, nil}, nil
}

func (s *AdodbStmt) Exec(args []driver.Value) (driver.Result, error) {
	list := make([]namedValue, len(args))
	for i, v := range args {
		list[i] = namedValue{
			Ordinal: i + 1,
			Value:   v,
		}
	}
	return s.exec(context.Background(), list)
}

func (s *AdodbStmt) exec(ctx context.Context, args []namedValue) (driver.Result, error) {
	if err := s.bind(args); err != nil {
		return nil, err
	}
	_, err := oleutil.CallMethod(s.s, "Execute")
	if err != nil {
		return nil, err
	}
	return driver.ResultNoRows, nil
}

type AdodbRows struct {
	s    *AdodbStmt
	rc   *ole.IDispatch
	nc   int
	cols []string
}

func (rc *AdodbRows) Close() error {
	rv, err := oleutil.CallMethod(rc.rc, "Close")
	if err != nil {
		return err
	}
	rv.Clear()
	rc.rc.Release()
	rc.rc = nil
	rc.s = nil
	return nil
}

func (rc *AdodbRows) Columns() []string {
	if rc.nc != len(rc.cols) {
		unknown, err := oleutil.GetProperty(rc.rc, "Fields")
		if err != nil {
			return []string{}
		}
		fields := unknown.ToIDispatch()
		unknown.Clear()
		defer fields.Release()
		rv, err := oleutil.GetProperty(fields, "Count")
		if err != nil {
			return []string{}
		}
		rc.nc = int(rv.Val)
		rv.Clear()
		rc.cols = make([]string, rc.nc)
		for i := 0; i < rc.nc; i++ {
			var varval ole.VARIANT
			varval.VT = ole.VT_I4
			varval.Val = int64(i)
			val, err := oleutil.CallMethod(fields, "Item", &varval)
			if err != nil {
				return []string{}
			}
			item := val.ToIDispatch()
			val.Clear()
			name, err := oleutil.GetProperty(item, "Name")
			if err != nil {
				item.Release()
				return []string{}
			}
			rc.cols[i] = name.ToString()
			name.Clear()
			item.Release()
		}
	}
	return rc.cols
}

func (rc *AdodbRows) Next(dest []driver.Value) error {
	eof, err := oleutil.GetProperty(rc.rc, "EOF")
	if err != nil {
		return io.EOF
	}
	if eof.Val != 0 {
		eof.Clear()
		return io.EOF
	}
	eof.Clear()

	unknown, err := oleutil.GetProperty(rc.rc, "Fields")
	if err != nil {
		return err
	}
	fields := unknown.ToIDispatch()
	unknown.Clear()
	defer fields.Release()
	for i := range dest {
		var varval ole.VARIANT
		varval.VT = ole.VT_I4
		varval.Val = int64(i)
		rv, err := oleutil.CallMethod(fields, "Item", &varval)
		if err != nil {
			return err
		}
		field := rv.ToIDispatch()
		rv.Clear()
		val, err := oleutil.GetProperty(field, "Value")
		if err != nil {
			field.Release()
			return err
		}
		if val.VT == 1 /* VT_NULL */ {
			dest[i] = nil
			val.Clear()
			field.Release()
			continue
		}
		typ, err := oleutil.GetProperty(field, "Type")
		if err != nil {
			val.Clear()
			field.Release()
			return err
		}
		sc, err := oleutil.GetProperty(field, "NumericScale")
		if err != nil {
			typ.Clear()
			val.Clear()
			field.Release()
			return err
		}
		switch typ.Val {
		case 0: // ADEMPTY
			dest[i] = nil
		case 2: // ADSMALLINT
			dest[i] = int64(int16(val.Val))
		case 3: // ADINTEGER
			dest[i] = int64(int32(val.Val))
		case 4: // ADSINGLE
			dest[i] = float64(math.Float32frombits(uint32(val.Val)))
		case 5: // ADDOUBLE
			dest[i] = math.Float64frombits(uint64(val.Val))
		case 6: // ADCURRENCY
			dest[i] = float64(val.Val) / 10000
		case 7: // ADDATE
			// see http://blogs.msdn.com/b/ericlippert/archive/2003/09/16/eric-s-complete-guide-to-vt-date.aspx
			d, t := math.Modf(math.Float64frombits(uint64(val.Val)))
			t = math.Abs(t)
			dest[i] = time.Date(1899, 12, 30+int(d), 0, 0, int(t*86400), 0, time.Local)
		case 8: // ADBSTR
			dest[i] = val.ToString()
		case 9: // ADIDISPATCH
			dest[i] = val.ToIDispatch()
		case 10: // ADERROR
			// TODO
		case 11: // ADBOOLEAN
			dest[i] = val.Val != 0
		case 12: // ADVARIANT
			dest[i] = val
		case 13: // ADIUNKNOWN
			dest[i] = val.ToIUnknown()
		case 14: // ADDECIMAL
			sub := math.Pow(10, float64(sc.Val))
			dest[i] = float64(float64(val.Val) / sub)
		case 16: // ADTINYINT
			dest[i] = int8(val.Val)
		case 17: // ADUNSIGNEDTINYINT
			dest[i] = uint8(val.Val)
		case 18: // ADUNSIGNEDSMALLINT
			dest[i] = uint16(val.Val)
		case 19: // ADUNSIGNEDINT
			dest[i] = uint32(val.Val)
		case 20: // ADBIGINT
			dest[i] = big.NewInt(val.Val)
		case 21: // ADUNSIGNEDBIGINT
			// TODO
		case 72: // ADGUID
			dest[i] = val.ToString()
		case 128: // ADBINARY
			sa := (*ole.SafeArray)(unsafe.Pointer(uintptr(val.Val)))
			conv := &ole.SafeArrayConversion{sa}
			elems, err := conv.TotalElements(0)
			if err != nil {
				return err
			}
			dest[i] = (*[1 << 30]byte)(unsafe.Pointer(uintptr(sa.Data)))[0:elems]
		case 129: // ADCHAR
			dest[i] = val.ToString() //uint8(val.Val)
		case 130: // ADWCHAR
			dest[i] = val.ToString() //uint16(val.Val)
		case 131: // ADNUMERIC
			sub := math.Pow(10, float64(sc.Val))
			dest[i] = float64(float64(val.Val) / sub)
		case 132: // ADUSERDEFINED
			dest[i] = uintptr(val.Val)
		case 133: // ADDBDATE
			// see http://blogs.msdn.com/b/ericlippert/archive/2003/09/16/eric-s-complete-guide-to-vt-date.aspx
			d := math.Float64frombits(uint64(val.Val))
			dest[i] = time.Date(1899, 12, 30+int(d), 0, 0, 0, 0, time.Local)
		case 134: // ADDBTIME
			t := math.Float64frombits(uint64(val.Val))
			dest[i] = time.Date(0, 1, 1, 0, 0, int(t*86400), 0, time.Local)
		case 135: // ADDBTIMESTAMP
			d, t := math.Modf(math.Float64frombits(uint64(val.Val)))
			t = math.Abs(t)
			dest[i] = time.Date(1899, 12, 30+int(d), 0, 0, int(t*86400), 0, time.Local)
		case 136: // ADCHAPTER
			dest[i] = val.ToString()
		case 200: // ADVARCHAR
			dest[i] = val.ToString()
		case 201: // ADLONGVARCHAR
			dest[i] = val.ToString()
		case 202: // ADVARWCHAR
			dest[i] = val.ToString()
		case 203: // ADLONGVARWCHAR
			dest[i] = val.ToString()
		case 204: // ADVARBINARY
			// TODO
		case 205: // ADLONGVARBINARY
			sa := (*ole.SafeArray)(unsafe.Pointer(uintptr(val.Val)))
			conv := &ole.SafeArrayConversion{sa}
			elems, err := conv.TotalElements(0)
			if err != nil {
				return err
			}
			dest[i] = (*[1 << 30]byte)(unsafe.Pointer(uintptr(sa.Data)))[0:elems]
		}
		typ.Clear()
		sc.Clear()
		val.Clear()
		field.Release()
	}
	rv, err := oleutil.CallMethod(rc.rc, "MoveNext")
	if err != nil {
		return err
	}
	rv.Clear()
	return nil
}

// ColumnTypeDatabaseTypeName implement RowsColumnTypeDatabaseTypeName.
func (rc *AdodbRows) ColumnTypeDatabaseTypeName(i int) string {
	if i >= rc.nc {
		return ""
	}
	unknown, err := oleutil.GetProperty(rc.rc, "Fields")
	if err != nil {
		return ""
	}
	fields := unknown.ToIDispatch()
	unknown.Clear()
	defer fields.Release()

	var varval ole.VARIANT
	varval.VT = ole.VT_I4
	varval.Val = int64(i)
	val, err := oleutil.CallMethod(fields, "Item", &varval)
	if err != nil {
		return ""
	}
	item := val.ToIDispatch()
	val.Clear()
	typ, err := oleutil.GetProperty(item, "Type")
	if err != nil {
		item.Release()
		return ""
	}
	typname := ""
	switch typ.Val {
	case 0:
		typname = "ADEMPTY"
	case 2:
		typname = "ADSMALLINT"
	case 3:
		typname = "ADINTEGER"
	case 4:
		typname = "ADSINGLE"
	case 5:
		typname = "ADDOUBLE"
	case 6:
		typname = "ADCURRENCY"
	case 7:
		typname = "ADDATE"
	case 8:
		typname = "ADBSTR"
	case 9:
		typname = "ADIDISPATCH"
	case 10:
		typname = "ADERROR"
	case 11:
		typname = "ADBOOLEAN"
	case 12:
		typname = "ADVARIANT"
	case 13:
		typname = "ADIUNKNOWN"
	case 14:
		typname = "ADDECIMAL"
	case 16:
		typname = "ADTINYINT"
	case 17:
		typname = "ADUNSIGNEDTINYINT"
	case 18:
		typname = "ADUNSIGNEDSMALLINT"
	case 19:
		typname = "ADUNSIGNEDINT"
	case 20:
		typname = "ADBIGINT"
	case 21:
		typname = "ADUNSIGNEDBIGINT"
	case 72:
		typname = "ADGUID"
	case 128:
		typname = "ADBINARY"
	case 129:
		typname = "ADCHAR"
	case 130:
		typname = "ADWCHAR"
	case 131:
		typname = "ADNUMERIC"
	case 132:
		typname = "ADUSERDEFINED"
	case 133:
		typname = "ADDBDATE"
	case 134:
		typname = "ADDBTIME"
	case 135:
		typname = "ADDBTIMESTAMP"
	case 136:
		typname = "ADCHAPTER"
	case 200:
		typname = "ADVARCHAR"
	case 201:
		typname = "ADLONGVARCHAR"
	case 202:
		typname = "ADVARWCHAR"
	case 203:
		typname = "ADLONGVARWCHAR"
	case 204:
		typname = "ADVARBINARY"
	case 205:
		typname = "ADLONGVARBINARY"
	}
	typ.Clear()
	item.Release()
	return typname
}

func (rc *AdodbRows) ColumnTypeLength(i int) (length int64, ok bool) {
	if i >= rc.nc {
		return 0, false
	}
	unknown, err := oleutil.GetProperty(rc.rc, "Fields")
	if err != nil {
		return 0, false
	}
	fields := unknown.ToIDispatch()
	unknown.Clear()
	defer fields.Release()

	var varval ole.VARIANT
	varval.VT = ole.VT_I4
	varval.Val = int64(i)
	val, err := oleutil.CallMethod(fields, "Item", &varval)
	if err != nil {
		return 0, false
	}
	item := val.ToIDispatch()
	val.Clear()
	siz, err := oleutil.GetProperty(item, "DefinedSize")
	if err != nil {
		item.Release()
		return 0, false
	}
	sizval := siz.Val
	siz.Clear()
	item.Release()
	return int64(sizval), true
}

// ColumnTypeNullable implement RowsColumnTypeNullable.
func (rc *AdodbRows) ColumnTypeNullable(i int) (nullable, ok bool) {
	if i >= rc.nc {
		return false, false
	}
	unknown, err := oleutil.GetProperty(rc.rc, "Fields")
	if err != nil {
		return false, false
	}
	fields := unknown.ToIDispatch()
	unknown.Clear()
	defer fields.Release()

	var varval ole.VARIANT
	varval.VT = ole.VT_I4
	varval.Val = int64(i)
	val, err := oleutil.CallMethod(fields, "Item", &varval)
	if err != nil {
		return false, false
	}
	item := val.ToIDispatch()
	val.Clear()
	att, err := oleutil.GetProperty(item, "Attributes")
	if err != nil {
		item.Release()
		return false, false
	}
	attributes := att.Val
	att.Clear()
	item.Release()
	return attributes&0x20 != 0, true
}

// ColumnTypeScanType implement RowsColumnTypeScanType.
func (rc *AdodbRows) ColumnTypeScanType(i int) reflect.Type {
	if i >= rc.nc {
		return reflect.TypeOf(nil)
	}
	unknown, err := oleutil.GetProperty(rc.rc, "Fields")
	if err != nil {
		return reflect.TypeOf(nil)
	}
	fields := unknown.ToIDispatch()
	unknown.Clear()
	defer fields.Release()

	var varval ole.VARIANT
	varval.VT = ole.VT_I4
	varval.Val = int64(i)
	val, err := oleutil.CallMethod(fields, "Item", &varval)
	if err != nil {
		return reflect.TypeOf(nil)
	}
	item := val.ToIDispatch()
	val.Clear()
	typ, err := oleutil.GetProperty(item, "Type")
	if err != nil {
		item.Release()
		return reflect.TypeOf(nil)
	}
	var rt reflect.Type
	switch typ.Val {
	case 0: // ADEMPTY
		rt = reflect.TypeOf(nil)
	case 2: // ADSMALLINT
		rt = reflect.TypeOf(int16(0))
	case 3: // ADINTEGER
		rt = reflect.TypeOf(int32(0))
	case 4: // ADSINGLE
		rt = reflect.TypeOf(float32(0))
	case 5: // ADDOUBLE
		rt = reflect.TypeOf(float64(0))
	case 6: // ADCURRENCY
		rt = reflect.TypeOf(float64(0))
	case 7: // ADDATE
		rt = reflect.TypeOf(time.Time{})
	case 8: // ADBSTR
		rt = reflect.TypeOf("")
	case 9: // ADIDISPATCH
		rt = reflect.TypeOf((*ole.IDispatch)(nil))
	case 10: // ADERROR
		rt = reflect.TypeOf((error)(nil))
	case 11: // ADBOOLEAN
		rt = reflect.TypeOf(true)
	case 12: // ADVARIANT
		var va ole.VARIANT
		rt = reflect.TypeOf(va)
	case 13: // ADIUNKNOWN
		rt = reflect.TypeOf((*ole.IUnknown)(nil))
	case 14: // ADDECIMAL
		rt = reflect.TypeOf(float64(0))
	case 16: // ADTINYINT
		rt = reflect.TypeOf(int8(0))
	case 17: // ADUNSIGNEDTINYINT
		rt = reflect.TypeOf(uint8(0))
	case 18: // ADUNSIGNEDSMALLINT
		rt = reflect.TypeOf(uint16(0))
	case 19: // ADUNSIGNEDINT
		rt = reflect.TypeOf(uint32(0))
	case 20: // ADBIGINT
		rt = reflect.TypeOf((*big.Int)(nil))
	case 21: // ADUNSIGNEDBIGINT
		rt = reflect.TypeOf(nil)
	case 72: // ADGUID
		var gi ole.GUID
		rt = reflect.TypeOf(gi)
	case 128: // ADBINARY
		rt = reflect.TypeOf((*ole.SafeArray)(nil))
	case 129: // ADCHAR
		rt = reflect.TypeOf(byte(0))
	case 130: // ADWCHAR
		rt = reflect.TypeOf(rune(0))
	case 131: // ADNUMERIC
		rt = reflect.TypeOf(float64(0))
	case 132: // ADUSERDEFINED
		rt = reflect.TypeOf(uintptr(0))
	case 133: // ADDBDATE
		rt = reflect.TypeOf(time.Time{})
	case 134: // ADDBTIME
		rt = reflect.TypeOf(time.Time{})
	case 135: // ADDBTIMESTAMP
		rt = reflect.TypeOf(time.Time{})
	case 136: // ADCHAPTER
		rt = reflect.TypeOf("")
	case 200: // ADVARCHAR
		rt = reflect.TypeOf("")
	case 201: // ADLONGVARCHAR
		rt = reflect.TypeOf("")
	case 202: // ADVARWCHAR
		rt = reflect.TypeOf("")
	case 203: // ADLONGVARWCHAR
		rt = reflect.TypeOf("")
	case 204: // ADVARBINARY
		rt = reflect.TypeOf([]byte{})
	case 205: // ADLONGVARBINARY
		rt = reflect.TypeOf((*ole.SafeArray)(nil))
	}
	typ.Clear()
	item.Release()
	return rt
}

func (rc *AdodbRows) ColumnTypePrecisionScale(i int) (precision, scale int64, ok bool) {
	if i >= rc.nc {
		return 0, 0, false
	}
	unknown, err := oleutil.GetProperty(rc.rc, "Fields")
	if err != nil {
		return 0, 0, false
	}
	fields := unknown.ToIDispatch()
	unknown.Clear()
	defer fields.Release()

	var varval ole.VARIANT
	varval.VT = ole.VT_I4
	varval.Val = int64(i)
	val, err := oleutil.CallMethod(fields, "Item", &varval)
	if err != nil {
		return 0, 0, false
	}

	item := val.ToIDispatch()
	val.Clear()

	typ, err := oleutil.GetProperty(item, "Type")
	if err != nil {
		item.Release()
		return 0, 0, false
	}
	if typ.Val != 131 {
		item.Release()
		return 0, 0, true
	}

	prec, err := oleutil.GetProperty(item, "Precision")
	if err != nil {
		item.Release()
		return 0, 0, false
	}

	scl, err := oleutil.GetProperty(item, "NumericScale")
	if err != nil {
		item.Release()
		return 0, 0, false
	}

	precval := prec.Val
	sclval := scl.Val
	prec.Clear()
	scl.Clear()
	item.Release()
	return int64(precval), int64(sclval), true
}
