
default: TurkCellTask.ts
	tsc $<
	lessc TurkCellTask.less > TurkCellTask.css

clean:
	rm *.js
