// Package dbaccess provides functions to access various databases
package dbaccess

import (
	"database/sql"
	"fmt"
	"github.com/go-ole/go-ole"
	_ "github.com/mattn/go-adodb"
)

// Msaccess provides connection to MS Access databases with the given path parameter.
// The path can be relative or absolute, e.g., "../some dir/file.accdb".
// It returns a db connection, a cleanup function to be deferred for resource release, and an error.
func Msaccess(path string) (*sql.DB, func(), error) {
	err := ole.CoInitialize(0)
	if err != nil {
		return nil, nil, fmt.Errorf("could not initialize COM library: %v", err)
	}

	cleanup := func() {
		ole.CoUninitialize()
	}

	provider := "Provider=Microsoft.ACE.OLEDB.12.0"
	db, err := sql.Open("adodb", fmt.Sprintf("%s;Data Source=%s;", provider, path))
	if err != nil {
		return nil, cleanup, fmt.Errorf("error while opening the database: %v", err)
	}

	return db, cleanup, nil
}

// ------------------------------------
// 	Msaccess provides connection to ms Access databases with path and table parameters:
// PARAMETERS: path, table name
//path: includes path to db file relative or absolute i.e. ../some dir/file.accdb;
//table: is the table name.
//	RETURNS: db connection string, rows in the table, and map of column names: types
//please note that you need 64 bit driver for 64 bit MS Access; drive bit = msaccess bit
//func Msaccess(path string, table string) (*sql.DB, *sql.Rows, map[string]string, error) {
//    err := ole.CoInitialize(0)
//    if err != nil {
//        return nil, nil, nil, fmt.Errorf("could not initialize COM library: %v", err)
//    }
//
//    provider := "Provider=Microsoft.ACE.OLEDB.12.0"
//    db, err := sql.Open("adodb", fmt.Sprintf("%s;Data Source=%s;", provider, path))
//    if err != nil {
//        ole.CoUninitialize()
//        return nil, nil, nil, fmt.Errorf("error while opening the database: %v", err)
//    }
//
//    rows, err := db.Query(fmt.Sprintf("SELECT * FROM %s", table))
//    if err != nil {
//        ole.CoUninitialize()
//        return nil, nil, nil, fmt.Errorf("error querying the table: %v", err)
//    }
//
//    columnTypes, err := rows.ColumnTypes()
//    if err != nil {
//        ole.CoUninitialize()
//        return nil, nil, nil, fmt.Errorf("error getting column types: %v", err)
//    }
//
//    tmap := make(map[string]string)
//    for _, colType := range columnTypes {
//        tmap[colType.Name()] = colType.DatabaseTypeName()
//    }
//
//    return db, rows, tmap, nil
//}
