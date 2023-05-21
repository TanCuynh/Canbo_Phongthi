import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.PreparedStatement;

public class canbo_phongthi {

    public static void main(String[] args) {
    	String jdbcUrl = "jdbc:mysql://localhost:3306/canbo_phongthi";
        String username = "root";
        String password = "";
        
        try (Connection connection = (Connection) DriverManager.getConnection(jdbcUrl, username, password);
                Workbook workbook = new XSSFWorkbook("F:\\WORKS\\Y3_ASSIGNMENTS\\TH_LTM\\DSCBCT.xlsx")) {

               for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                   Sheet sheet = workbook.getSheetAt(sheetIndex);

                   if (sheetIndex == 0) {
                       saveSheetDataToTable1(connection, sheet);
                   } else if (sheetIndex == 1) {
                       saveSheetDataToTable2(connection, sheet);
                   }
               }

               System.out.println("Doc du lieu vao Database thanh cong");

           } catch (SQLException | IOException e) {
               e.printStackTrace();
           }
        
        try {
            // Đọc dữ liệu từ file Excel danh sách cán bộ và danh sách phòng thi
            List<CanBo> danhSachCanBo = docDanhSachCanBoTuExcel("F:\\WORKS\\Y3_ASSIGNMENTS\\TH_LTM\\DSCBCT.xlsx");
            List<PhongThi> danhSachPhongThi = docDanhSachPhongThiTuExcel("F:\\WORKS\\Y3_ASSIGNMENTS\\TH_LTM\\DSCBCT.xlsx");
           
            // Xáo trộn danh sách cán bộ
            Collections.shuffle(danhSachCanBo);

            // Tạo workbook mới để ghi kết quả
            Workbook workbook = new XSSFWorkbook();
            Sheet danhSachPhanCong = workbook.createSheet("DANHSACH PHANCONG");
            Sheet danhSachGiamSat = workbook.createSheet("DANHSACH GIAMSAT");
            
            // Tạo hàng tiêu đề
            Row headerRow = danhSachPhanCong.createRow(0);
            headerRow.createCell(0).setCellValue("STT");
            headerRow.createCell(1).setCellValue("Mã GV");
            headerRow.createCell(2).setCellValue("Họ Tên");
            headerRow.createCell(3).setCellValue("Giám thị 1");
            headerRow.createCell(4).setCellValue("Giám thị 2");
            headerRow.createCell(5).setCellValue("Mã phòng");
            
            Row headerRowGiamSat = danhSachGiamSat.createRow(0);
            headerRowGiamSat.createCell(0).setCellValue("STT");
            headerRowGiamSat.createCell(1).setCellValue("Mã GV");
            headerRowGiamSat.createCell(2).setCellValue("Họ Tên");
            headerRowGiamSat.createCell(3).setCellValue("Phòng thi được giám sát");
            

            int rowIndex = 1; // Bắt đầu ghi từ hàng thứ 2
            int stt = 1;	// số thứ tự

            // Phân bố cán bộ vào phòng thi
            for (PhongThi phongThi : danhSachPhongThi) {
                for (int i = 0; i < 2; i++) {
                    if (danhSachCanBo.isEmpty()) {
                        break; // Nếu danh sách cán bộ đã hết, thoát khỏi vòng lặp
                    }

                    CanBo canBo = danhSachCanBo.remove(0); // Lấy cán bộ từ danh sách

                    Row dataRow = danhSachPhanCong.createRow(rowIndex);
                    dataRow.createCell(0).setCellValue(stt);
                    dataRow.createCell(1).setCellValue(canBo.getMaGV());
                    dataRow.createCell(2).setCellValue(canBo.getHoTen());
                    dataRow.createCell(3).setCellValue(i == 0 ? "X" : ""); // Giám thị 1
                    dataRow.createCell(4).setCellValue(i == 1 ? "X" : ""); // Giám thị 2 
                    dataRow.createCell(5).setCellValue(phongThi.getMaPhong());
                    rowIndex++;
                    stt++;
                }
            }
            
            int rowIndexGiamSat = 1;
            int sttGiamSat = 1;
           
            if (!danhSachCanBo.isEmpty()) {
                for (CanBo canBo : danhSachCanBo) {
                	
                	if (danhSachPhongThi.isEmpty()) {
                        break; // Nếu danh sách phong thi đã hết, thoát khỏi vòng lặp
                    }
                	
                		int remainingRooms = danhSachPhongThi.size();
                	
                		PhongThi phongThi =  danhSachPhongThi.get(0);
                		for (int i =0 ;i < Math.min(remainingRooms, 10);i++) {
                    		danhSachPhongThi.remove(0);

                		}
                	
                    int startRoom = phongThi.getMaPhong(); // Tính phòng đầu tiên của range
                    int endRoom = startRoom + Math.min(remainingRooms - 1, 9); // Tính phòng cuối cùng của range
                    String supervisionRoomRange = "Từ " + startRoom + " đến " + endRoom;
                	
                	
                	
                    Row dataRow = danhSachGiamSat.createRow(rowIndexGiamSat);
                    dataRow.createCell(0).setCellValue(sttGiamSat); // STT
                    dataRow.createCell(1).setCellValue(canBo.getMaGV());
                    dataRow.createCell(2).setCellValue(canBo.getHoTen());
                    dataRow.createCell(3).setCellValue(supervisionRoomRange); // Phòng thi được giám sát
                    rowIndexGiamSat++;
                    sttGiamSat++;
                }
            }
                        
            FileOutputStream outputStream = new FileOutputStream("F:\\WORKS\\Y3_ASSIGNMENTS\\TH_LTM\\Danh sach phan cong va giam sat.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println("Ket qua da duoc ghi vao sheet Danh sach phan cong va giam sat");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    private static void saveSheetDataToTable1(Connection connection, Sheet sheet) throws SQLException {
        String sql = "INSERT INTO canbo (stt, MaGV, HoTen, NgaySinh, DonViCongTac) VALUES (?, ?, ?, ?, ?)";
        PreparedStatement statement = (PreparedStatement) connection.prepareStatement(sql);
        
        
        
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
        	
            Row row = sheet.getRow(rowIndex);
            int cellCount = row.getLastCellNum();
            int rowisNull = 1;
            
        	for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
    	        Cell cell = row.getCell(cellIndex);

    	        if (cell != null && cell.getCellType() != CellType.BLANK) {
    	        	rowisNull = 0;
    	            break;
    	        }
    	    }
    	    if(rowisNull == 1)
    	    {
    	    	break;
    	    }
            
            Cell cell1 = row.getCell(0);
            Cell cell2 = row.getCell(1);
            Cell cell3 = row.getCell(2);
            Cell cell4 = row.getCell(3);
            Cell cell5 = row.getCell(4);

            statement.setDouble(1, cell1.getNumericCellValue());
            statement.setDouble(2, cell2.getNumericCellValue());
            statement.setString(3, cell3.getStringCellValue());
            statement.setDate(4, new java.sql.Date(cell4.getDateCellValue().getTime()));
            statement.setString(5, cell5.getStringCellValue());

            statement.executeUpdate();
        }

        statement.close();
    }

    private static void saveSheetDataToTable2(Connection connection, Sheet sheet) throws SQLException {
        String sql = "INSERT INTO phongthi (stt, PhongThi, GhiChu) VALUES (?, ?, ?)";
        PreparedStatement statement = (PreparedStatement) connection.prepareStatement(sql);
        
        
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
        	Row row = sheet.getRow(rowIndex);
        	int cellCount = row.getLastCellNum();
        	int rowisNull = 1;
        	
        	for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
    	        Cell cell = row.getCell(cellIndex);

    	        if (cell != null && cell.getCellType() != CellType.BLANK) {
    	        	rowisNull = 0;
    	            break;
    	        }
    	    }
    	    if(rowisNull == 1)
    	    {
    	    	break;
    	    }
        	
            Cell cell1 = row.getCell(0);
            Cell cell2 = row.getCell(1);
            Cell cell3 = row.getCell(2);

            statement.setDouble(1, cell1.getNumericCellValue());
            statement.setInt(2,(int) cell2.getNumericCellValue());
            statement.setString(3, cell3.getStringCellValue());

            statement.executeUpdate();
        }

        statement.close();
    }

    // Đọc danh sách cán bộ từ file Excel
    private static List<CanBo> docDanhSachCanBoTuExcel(String filePath) throws FileNotFoundException, IOException {
        List<CanBo> danhSachCanBo = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {
	        	Sheet canboSheet = workbook.getSheetAt(0);
	        	int rowCount = canboSheet.getLastRowNum();
	        	int cellCount = canboSheet.getRow(rowCount).getLastCellNum();
	        	
	        	
	        	for (int rowIndex = 0; rowIndex <= rowCount; rowIndex++) {
	        	    Row row = canboSheet.getRow(rowIndex);
	        	    int rowisNull = 1;
	        	    if (rowIndex == 0) {
	        	        continue;
	        	    }
	        	    for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
	        	        Cell cell = row.getCell(cellIndex);

	        	        if (cell != null && cell.getCellType() != CellType.BLANK) {
	        	        	rowisNull = 0;
	        	            break;
	        	        }
	        	    }
	        	    if(rowisNull == 1)
	        	    {
	        	    	break;
	        	    }
	        	    int maGV =(int) row.getCell(1).getNumericCellValue();
	                String hoTen = row.getCell(2).getStringCellValue();
	                danhSachCanBo.add(new CanBo(maGV, hoTen));
	        	}
        }

        return danhSachCanBo;
    }

    // Đọc danh sách phòng thi từ file Excel
    private static List<PhongThi> docDanhSachPhongThiTuExcel(String filePath) throws FileNotFoundException, IOException {
    	List<PhongThi> danhSachPhongThi = new ArrayList<>();
    	
    	try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
    		
    		Sheet canboSheet = workbook.getSheetAt(1);
        	int rowCount = canboSheet.getLastRowNum();
        	int cellCount = canboSheet.getRow(rowCount).getLastCellNum();
        	
        	
        	for (int rowIndex = 0; rowIndex <= rowCount; rowIndex++) {
        	    Row row = canboSheet.getRow(rowIndex);
        	    int rowisNull = 1;
        	    if (rowIndex == 0) {
        	        continue;
        	    }
        	    for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
        	        Cell cell = row.getCell(cellIndex);

        	        if (cell != null && cell.getCellType() != CellType.BLANK) {
        	        	rowisNull = 0;
        	            break;
        	        }
        	    }
        	    if(rowisNull == 1)
        	    {
        	    	break;
        	    }
        	    int maPhongThi = (int) row.getCell(1).getNumericCellValue();
                danhSachPhongThi.add(new PhongThi(maPhongThi));
        	}
    	}

        return danhSachPhongThi;
    }

    // Lớp đại diện cho cán bộ
    private static class CanBo {
        private int maGV;
        private String hoTen;

        public CanBo(int maGV, String hoTen) {
            this.maGV = maGV;
            this.hoTen = hoTen;
        }

        public int getMaGV() {
            return maGV;
        }

        public String getHoTen() {
            return hoTen;
        }
    }

    // Lớp đại diện cho phòng thi
    private static class PhongThi {
        private int maPhong;

        public PhongThi(int maPhong) {
            this.maPhong = maPhong;
        }

        public int getMaPhong() {
            return maPhong;
        }
    }
}