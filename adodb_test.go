package adodb

import (
	"testing"
)

import (
	"database/sql"
	"fmt"
	"os"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

var provider string

func createMdb(f string) error {
	ole.CoInitialize(0)

	unk, err := oleutil.CreateObject("ADOX.Catalog")
	if err != nil {
		return err
	}
	defer unk.Release()
	cat, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer cat.Release()
	provider = "Microsoft.Jet.OLEDB.4.0"
	r, err := oleutil.CallMethod(cat, "Create", "Provider="+provider+";Data Source="+f+";")
	if err != nil {
		provider = "Microsoft.ACE.OLEDB.12.0"
		r, err = oleutil.CallMethod(cat, "Create", "Provider="+provider+";Data Source="+f+";")
		if err != nil {
			return err
		}
	}
	r.Clear()
	return nil
}

func TestSimple(t *testing.T) {
	f := "./example.mdb"

	os.Remove(f)

	err := createMdb(f)
	if err != nil {
		t.Fatal(err)
	}

	db, err := sql.Open("adodb", "Provider="+provider+";Data Source="+f+";")
	if err != nil {
		t.Fatal(err)
	}
	defer db.Close()

	_, err = db.Exec("create table foo (id int not null primary key, name text not null, created datetime not null)")
	if err != nil {
		t.Fatal(err)
	}

	tx, err := db.Begin()
	if err != nil {
		t.Fatal(err)
	}
	stmt, err := tx.Prepare("insert into foo(id, name, created) values(?, ?, ?)")
	if err != nil {
		t.Fatal(err)
	}
	defer stmt.Close()

	for i := 0; i < 1000; i++ {
		_, err = stmt.Exec(i, fmt.Sprintf("こんにちわ世界%03d", i), time.Now())
		if err != nil {
			t.Fatal(err)
		}
	}
	tx.Commit()

	rows, err := db.Query("select id, name, created from foo")
	if err != nil {
		fmt.Println("select", err)
		return
	}
	defer rows.Close()

	for rows.Next() {
		var id int
		var name string
		var created time.Time
		err = rows.Scan(&id, &name, &created)
		if err != nil {
			t.Fatal(err)
		}
		fmt.Println(id, name, created)
	}
}
