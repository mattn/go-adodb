package main

import (
	"database/sql"
	"fmt"
	"os"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	_ "github.com/mattn/go-adodb"
)

func createMdb(f string) error {
	unk, err := oleutil.CreateObject("ADOX.Catalog")
	if err != nil {
		return err
	}
	cat, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	_, err = oleutil.CallMethod(cat, "Create", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+f+";")
	if err != nil {
		return err
	}
	return nil
}

func main() {
	ole.CoInitialize(0)

	f := "./example.mdb"

	os.Remove(f)

	err := createMdb(f)
	if err != nil {
		fmt.Println("create mdb", err)
		return
	}

	db, err := sql.Open("adodb", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+f+";")
	if err != nil {
		fmt.Println("open", err)
		return
	}

	_, err = db.Exec("create table foo (id int not null primary key, name text not null, created datetime not null)")
	if err != nil {
		fmt.Println("create table", err)
		return
	}

	tx, err := db.Begin()
	if err != nil {
		fmt.Println(err)
		return
	}
	stmt, err := tx.Prepare("insert into foo(id, name, created) values(?, ?, ?)")
	if err != nil {
		fmt.Println("insert", err)
		return
	}
	defer stmt.Close()

	for i := 0; i < 100; i++ {
		_, err = stmt.Exec(i, fmt.Sprintf("こんにちわ世界%03d", i), time.Now())
		if err != nil {
			fmt.Println("exec", err)
			return
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
			fmt.Println("scan", err)
			return
		}
		fmt.Println(id, name, created)
	}
}
