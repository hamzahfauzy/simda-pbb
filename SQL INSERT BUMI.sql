CREATE PROC INSERT_BUMI
@xProp nVarchar(3), @xKab nVarchar(3), @xxKec nVarchar(3), @xxKel nVarchar(3),
 @xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1), @xNo smallint, 
 @xZNT  nVarchar(2), @Luas bigint , @xJBUMI nVarchar(1) , @xNilai bigint, @xForm nVarchar(11),
 @xStatus int,
 @xProp1 nVarchar(3), @xKab1 nVarchar(3), @xxKec1 nVarchar(3), @xxKel1 nVarchar(3),
 @xxBlok1 nVarchar(3), @xxUrut1 nVarchar(4), @xxJenis1 nVarchar(1),
 @xID nvarchar(30),@xForm1 nVarchar(11),@xPersil nvarchar(5),@xJalanOP nvarchar(30),@xBlokOP nvarchar(15),
 @xRW nvarchar(2), @xRT nvarchar(2),@xStatusWP nvarchar(1),@xLBumi bigint, @xNJOP_BM bigint,
 @xTrans nvarchar(1),@xTgl1 datetime, @xNIP1 nvarchar(30),@xTgl2 datetime, @xNIP2 nvarchar(30),
 @xTgl3 datetime, @xNIP3 nvarchar(30),@xCabang smallint,@xLBangunan bigint,@xNJOP_BG bigint,@xPeta smallint
AS
BEGIN
insert into DAT_OP_BUMI(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,
			NO_BUMI,KD_ZNT,LUAS_BUMI,JNS_BUMI, NILAI_SISTEM_BUMI,NO_FORMULIR,STATUS_JADI )
Values(@xProp, @xKab, @xxKec, @xxKel, @xxBlok, @xxUrut, @xxJenis, @xNo, @xZNT, @Luas, @xJBUMI, @xNilai, @xForm, @xStatus)
INSERT INTO DAT_OBJEK_PAJAK(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,
			SUBJEK_PAJAK_ID,NO_FORMULIR_SPOP,NO_PERSIL,JALAN_OP, BLOK_KAV_NO_OP,RW_OP,RT_OP,KD_STATUS_WP,
			TOTAL_LUAS_BUMI,NJOP_BUMI,JNS_TRANSAKSI_OP,TGL_PENDATAAN_OP,NIP_PENDATA,TGL_PEMERIKSAAN_OP,
			NIP_PEMERIKSA_OP,TGL_PEREKAMAN_OP,NIP_PEREKAM_OP,KD_STATUS_CABANG,TOTAL_LUAS_BNG,NJOP_BNG,STATUS_PETA_OP)
	Values(@xProp1, @xKab1, @xxKec1, @xxKel1, @xxBlok1, @xxUrut1, @xxJenis1, 
		   @xID, @xForm1, @xPersil, @xJalanOP , @xBlokOP , @xRW, @xRT,@xStatusWP,
		   @xLBumi, @xNJOP_BM, @xTrans,@xTgl1, @xNIP1,@xTgl2, 
		   @xNIP2, @xTgl3, @xNIP3,@xCabang,@xLBangunan,@xNJOP_BG,@xPeta)
IF @@ERROR<>0 
BEGIN
ROLLBACK TRAN
END
ELSE
BEGIN
COMMIT TRAN
END
END