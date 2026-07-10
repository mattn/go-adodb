package adodb

import (
	"bytes"
	"database/sql"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"time"

	// Importing runtime/cgo expands the initial goroutine stack, which ADO
	// needs to run reliably. See the note in README.md.
	_ "runtime/cgo"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// createDB creates an empty Access database and returns a DSN for it, or skips
// the test when no usable ADODB provider (ACE/Jet) is installed.
func createDB(t *testing.T) string {
	t.Helper()

	// COM must be initialized on the OS thread that runs the test, so pin the
	// test goroutine to its thread for the whole test.
	runtime.LockOSThread()
	ole.CoInitialize(0)

	path := filepath.Join(t.TempDir(), "test.mdb")
	os.Remove(path)

	unk, err := oleutil.CreateObject("ADOX.Catalog")
	if err != nil {
		t.Skipf("ADOX.Catalog unavailable, skipping: %v", err)
	}
	defer unk.Release()
	cat, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		t.Fatal(err)
	}
	defer cat.Release()

	for _, provider := range []string{"Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0"} {
		r, err := oleutil.CallMethod(cat, "Create", "Provider="+provider+";Data Source="+path+";")
		if err == nil {
			r.Clear()
			return "Provider=" + provider + ";Data Source=" + path + ";"
		}
	}
	t.Skip("no usable ADODB provider (ACE/Jet) installed, skipping")
	return ""
}

func TestQuery(t *testing.T) {
	dsn := createDB(t)
	db, err := sql.Open("adodb", dsn)
	if err != nil {
		t.Fatal(err)
	}
	defer db.Close()

	if _, err = db.Exec("create table foo (id int not null primary key, name text not null, created datetime not null)"); err != nil {
		t.Fatal(err)
	}

	const n = 100
	base := time.Date(2020, 1, 1, 0, 0, 0, 0, time.Local)
	stmt, err := db.Prepare("insert into foo(id, name, created) values(?, ?, ?)")
	if err != nil {
		t.Fatal(err)
	}
	for i := 0; i < n; i++ {
		if _, err = stmt.Exec(i, fmt.Sprintf("name%03d", i), base.Add(time.Duration(i)*time.Hour)); err != nil {
			t.Fatal(err)
		}
	}
	stmt.Close()

	// Run the query many times to shake out reference-counting problems that
	// only appear after repeated use (see issue #28).
	for iter := 0; iter < 50; iter++ {
		rows, err := db.Query("select id, name, created from foo order by id")
		if err != nil {
			t.Fatal(err)
		}
		count := 0
		for rows.Next() {
			var id int
			var name string
			var created time.Time
			if err = rows.Scan(&id, &name, &created); err != nil {
				t.Fatal(err)
			}
			if want := fmt.Sprintf("name%03d", id); name != want {
				t.Fatalf("iter %d: id %d: name = %q, want %q", iter, id, name, want)
			}
			count++
		}
		if err = rows.Err(); err != nil {
			t.Fatal(err)
		}
		rows.Close()
		if count != n {
			t.Fatalf("iter %d: got %d rows, want %d", iter, count, n)
		}
	}
}

func TestBinary(t *testing.T) {
	dsn := createDB(t)
	db, err := sql.Open("adodb", dsn)
	if err != nil {
		t.Fatal(err)
	}
	defer db.Close()

	if _, err = db.Exec("create table blobs (id int not null primary key, data image)"); err != nil {
		t.Fatal(err)
	}

	want := []byte{0x00, 0x01, 0x02, 0xfd, 0xfe, 0xff, 'g', 'o'}
	if _, err = db.Exec("insert into blobs(id, data) values(?, ?)", 1, want); err != nil {
		t.Fatal(err)
	}

	// Query repeatedly; the binary value must round-trip byte-for-byte and must
	// not be read from freed SafeArray memory (issue fixed in #62).
	for iter := 0; iter < 20; iter++ {
		var got []byte
		if err = db.QueryRow("select data from blobs where id = 1").Scan(&got); err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(got, want) {
			t.Fatalf("iter %d: data = %v, want %v", iter, got, want)
		}
	}
}
