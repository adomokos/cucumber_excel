class ExcelHandler
  include Singleton
	
  # set this to your Excel file path
  @@excel_file_path = 'C:\Temp\TestWorkbook.xlsx'

  def open_excel
    begin
      @excel = WIN32OLE.connect('excel.application')			
      @wb = @excel.ActiveWorkbook
    rescue
      @excel = WIN32OLE::new("excel.application")				
      @excel.visible =true
	  @wb = @excel.Workbooks.Open(@@excel_file_path )
	end	
  end
	
  def worksheet
    @wb.worksheets(1)
  end
	
  def close_excel    
    kill_excel
  end
	
private
  def kill_excel	
    wmi = WIN32OLE.connect("winmgmts://")
    processes = wmi.ExecQuery("select * from win32_process where
   commandline like '%excel.exe\"% /automation %'")
    for process in processes do
      Process.kill( 'KILL', process.ProcessID.to_i)
    end
  end
end