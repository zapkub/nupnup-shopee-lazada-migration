package main

import (
	"log"
	"os"
	"path"

	"github.com/zapkub/nupnup-shopee-migration/x/pipeline"
)

func main() {

	var wd, _ = os.Getwd()
	var p = pipeline.New()
	log.SetFlags(log.Lmsgprefix)
	var err error
	err = p.LoadExcludeProductID([]string{
		path.Join(wd, "datasource/archive-2021-07/basic-info-001-bk.xlsx"),
	})
	if err != nil {
		panic(err)
	}

	err = p.LoadSaleInfoFiles([]string{
		path.Join(wd, "datasource/sales-info-001.xlsx"),
	})

	if err != nil {
		panic(err)
	}

	err = p.LoadMediaInfoFile([]string{
		path.Join(wd, "datasource/media-info-001.xlsx"),
	})
	if err != nil {
		panic(err)
	}

	err = p.LoadBasicInfoFile([]string{
		path.Join(wd, "datasource/basic-info-001.xlsx"),
	})
	if err != nil {
		panic(err)
	}

	err = p.LoadPriceInfoFiles([]string{
		path.Join(wd, "datasource/discount_nominate_1000-1.xlsx"),
		path.Join(wd, "datasource/discount_nominate_3000-1.xlsx"),
		path.Join(wd, "datasource/discount_nominate_3000-2.xlsx"),
		path.Join(wd, "datasource/discount_nominate_3000-3.xlsx"),
		path.Join(wd, "datasource/discount_nominate_5000-2.xlsx"),
	})
	if err != nil {
		panic(err)
	}

	err = p.Process()
	if err != nil {
		panic(err)
	}

	outputFile, err := os.OpenFile(path.Join(wd, "output.xlsx"), os.O_RDWR|os.O_CREATE, 0655)
	if err != nil {
		panic(err)
	}
	p.Write(outputFile)

}
