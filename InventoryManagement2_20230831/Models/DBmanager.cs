using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace InventoryManagement3.Models
{
    public class DBmanager
    {
        private readonly string ConnStr = "Data Source=X-LAURENLO-23;Initial Catalog=InventoryDemo;Integrated Security=True";

        public List<Place> GetPlaces()
        {
            List<Place> place_list = new List<Place>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM place ORDER BY PlaceID ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    Place place = new Place
                    {
                        PlaceID = reader.GetString(reader.GetOrdinal("PlaceID")),
                        PlaceName = reader.GetString(reader.GetOrdinal("PlaceName"))
                    };
                    place_list.Add(place);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return place_list;
        }

        public void NewPlace(Place place_list)
        {
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO place (PlaceID, PlaceName) VALUES (@PlaceID, @PlaceName)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceID", place_list.PlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceName", place_list.PlaceName));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }

        public Place GetPlaceByPlaceID(string PlaceID)
        {
            Place place = new Place();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM place WHERE PlaceID = @PlaceID");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceID",PlaceID));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    place = new Place
                    {
                        PlaceID = reader.GetString(reader.GetOrdinal("PlaceID")),
                        PlaceName = reader.GetString(reader.GetOrdinal("PlaceName")),
                    };
                }
            }
            else
            {
                place.PlaceID = "未找到該筆資料";
            }
            sqlConnection.Close();
            return place;
        }

        public void UpdatePlace(Place place)
        {
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE place SET PlaceID = @PlaceID, PlaceName = @PlaceName WHERE PlaceID = @PlaceID");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceID", place.PlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceName", place.PlaceName));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }

        public void DeletePlaceByPlaceID(string PlaceID)
        {
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM place WHERE PlaceID = @PlaceID");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceID", PlaceID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }

        public List<GoodsReceipt> GetGoodsReceipts()
        {
            List<GoodsReceipt> goods_receipt = new List<GoodsReceipt>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM goodsreceipt ORDER BY GRSN ASC, Item ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    GoodsReceipt goodsreceipt = new GoodsReceipt
                    {
                        GRSN = reader.GetString(reader.GetOrdinal("GRSN")),
                        ReceiptDTM = reader.GetDateTime(reader.GetOrdinal("ReceiptDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        GRQty = reader.GetInt32(reader.GetOrdinal("GRQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                    goods_receipt.Add(goodsreceipt);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return goods_receipt;
        }
        public void NewGoodsReceipt(GoodsReceipt goodsreceipt)
        {
            //建立進貨單
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO goodsreceipt (GRSN, ReceiptDTM, Item, PartId, GRQty, SupplierID, MoveInPlaceID, EmployeeID)
                VALUES (@GRSN, @ReceiptDTM, @Item, @PartID, @GRQty, @SupplierID, @MoveInPlaceID, @EmployeeID)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@GRSN", goodsreceipt.GRSN));
            DateTime date = DateTime.Now;
            sqlCommand.Parameters.Add(new SqlParameter("@ReceiptDTM", date));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", goodsreceipt.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@GRQty", goodsreceipt.GRQty));
            sqlCommand.Parameters.Add(new SqlParameter("@SupplierID", goodsreceipt.SupplierID));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveInPlaceID", goodsreceipt.MoveInPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", goodsreceipt.EmployeeID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", goodsreceipt.GRSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", goodsreceipt.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", goodsreceipt.MoveInPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", goodsreceipt.GRQty));
            int qty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT GRQty from goodsreceipt WHERE GRSN = @GRSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@GRSN", goodsreceipt.GRSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", goodsreceipt.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", goodsreceipt.MoveInPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public GoodsReceipt GetGoodsReceiptByGRSN(string GRSN, int Item)
        {
            GoodsReceipt goodsreceipt = new GoodsReceipt();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM goodsreceipt WHERE GRSN = @GRSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@GRSN", GRSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    goodsreceipt = new GoodsReceipt
                    {
                        GRSN = reader.GetString(reader.GetOrdinal("GRSN")),
                        ReceiptDTM = reader.GetDateTime(reader.GetOrdinal("ReceiptDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        GRQty = reader.GetInt32(reader.GetOrdinal("GRQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            else
            {
                goodsreceipt.GRSN = "未找到該筆資料";
            }
            sqlConnection.Close();
            return goodsreceipt;
        }

        public void UpdateGoodsReceipt(GoodsReceipt goodsreceipt)
        {
            GoodsReceipt ReverseGR = new GoodsReceipt();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseGR = new SqlCommand("SELECT * FROM goodsreceipt WHERE GRSN = @GRSN AND Item = @Item");
            sqlCommandReverseGR.Connection = sqlConnection;
            sqlCommandReverseGR.Parameters.Add(new SqlParameter("@GRSN", goodsreceipt.GRSN));
            sqlCommandReverseGR.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseGR.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseGR = new GoodsReceipt
                    {
                        GRSN = reader.GetString(reader.GetOrdinal("GRSN")),
                        ReceiptDTM = reader.GetDateTime(reader.GetOrdinal("ReceiptDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        GRQty = reader.GetInt32(reader.GetOrdinal("GRQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseGR.GRSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseGR.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseGR.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ReverseGR.MoveInPlaceID));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));            
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", ReverseGR.GRQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT GRQty from goodsreceipt WHERE GRSN = @GRSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@GRSN", goodsreceipt.GRSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseGR.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseGR.MoveInPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //修改進貨單
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE goodsreceipt SET GRSN = @GRSN, ReceiptDTM = @ReceiptDTM, Item = @Item, PartID = @PartID, GRQty = @GRQty, 
                SupplierID = @SupplierID, MoveInPlaceID = @MoveInPlaceID, EmployeeID = @EmployeeID WHERE GRSN = @GRSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@GRSN", goodsreceipt.GRSN));
            sqlCommand.Parameters.Add(new SqlParameter("@ReceiptDTM", goodsreceipt.ReceiptDTM));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", goodsreceipt.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@GRQty", goodsreceipt.GRQty));
            sqlCommand.Parameters.Add(new SqlParameter("@SupplierID", goodsreceipt.SupplierID));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveInPlaceID", goodsreceipt.MoveInPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", goodsreceipt.EmployeeID));           
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", goodsreceipt.GRSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", goodsreceipt.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", goodsreceipt.MoveInPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", goodsreceipt.GRQty));
            int StockChangeQty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", StockChangeQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand SqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT GRQty from goodsreceipt WHERE GRSN = @GRSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            SqlCommandStock.Connection = sqlConnection;
            SqlCommandStock.Parameters.Add(new SqlParameter("@GRSN", goodsreceipt.GRSN));
            SqlCommandStock.Parameters.Add(new SqlParameter("@Item", goodsreceipt.Item));
            SqlCommandStock.Parameters.Add(new SqlParameter("@PartID", goodsreceipt.PartID));            
            SqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", goodsreceipt.MoveInPlaceID));
            SqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();

        }

        public void DeleteGoodsReceiptByGRSN(string GRSN, int Item)
        {
            GoodsReceipt ReverseGR = new GoodsReceipt();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseGR = new SqlCommand("SELECT * FROM goodsreceipt WHERE GRSN = @GRSN AND Item = @Item");
            sqlCommandReverseGR.Connection = sqlConnection;
            sqlCommandReverseGR.Parameters.Add(new SqlParameter("@GRSN", GRSN));
            sqlCommandReverseGR.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseGR.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseGR = new GoodsReceipt
                    {
                        GRSN = reader.GetString(reader.GetOrdinal("GRSN")),
                        ReceiptDTM = reader.GetDateTime(reader.GetOrdinal("ReceiptDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        GRQty = reader.GetInt32(reader.GetOrdinal("GRQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseGR.GRSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseGR.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseGR.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ReverseGR.MoveInPlaceID));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", ReverseGR.GRQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            
            //沖正庫存
            SqlCommand SqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT GRQty from goodsreceipt WHERE GRSN = @GRSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            SqlCommandStock.Connection = sqlConnection;
            SqlCommandStock.Parameters.Add(new SqlParameter("@GRSN", GRSN));
            SqlCommandStock.Parameters.Add(new SqlParameter("@Item", Item));
            SqlCommandStock.Parameters.Add(new SqlParameter("@PartID", ReverseGR.PartID));
            SqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseGR.MoveInPlaceID)); 
            SqlCommandStock.ExecuteNonQuery();

            //刪除進貨單
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM goodsreceipt WHERE GRSN = @GRSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@GRSN", GRSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));            
            sqlCommand.ExecuteNonQuery();            

            sqlConnection.Close();
        }

        public List<Returns> GetReturns()
        {
            List<Returns> returns_list = new List<Returns>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM returns ORDER BY ReturnsSN ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    Returns returns = new Returns
                    {
                        ReturnsSN = reader.GetString(reader.GetOrdinal("ReturnsSN")),
                        ReturnsDTM = reader.GetDateTime(reader.GetOrdinal("ReturnsDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ReturnsQty = reader.GetInt32(reader.GetOrdinal("ReturnsQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID")),
                    };
                    returns_list.Add(returns);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return returns_list;
        }

        public void NewReturns(Returns returns)
        {
            //建立退貨單
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO returns (ReturnsSN, ReturnsDTM, Item, PartId, ReturnsQty, SupplierID, MoveOutPlaceID, EmployeeID)
                VALUES (@ReturnsSN, @ReturnsDTM, @Item, @PartID, @ReturnsQty, @SupplierID, @MoveOutPlaceID, @EmployeeID)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsSN", returns.ReturnsSN));
            DateTime date = DateTime.Now;
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsDTM", date));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", returns.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsQty", returns.ReturnsQty));
            sqlCommand.Parameters.Add(new SqlParameter("@SupplierID", returns.SupplierID));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", returns.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", returns.EmployeeID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", returns.ReturnsSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", returns.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", returns.MoveOutPlaceID));
            int qty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", returns.ReturnsQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT ReturnsQty from returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID
                ");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@ReturnsSN", returns.ReturnsSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", returns.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", returns.MoveOutPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public Returns GetReturnsByReturnsSN(string ReturnsSN, int Item)
        {
            Returns returns = new Returns();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsSN", ReturnsSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    returns = new Returns
                    {
                        ReturnsSN = reader.GetString(reader.GetOrdinal("ReturnsSN")),
                        ReturnsDTM = reader.GetDateTime(reader.GetOrdinal("ReturnsDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ReturnsQty = reader.GetInt32(reader.GetOrdinal("ReturnsQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            else
            {
                returns.ReturnsSN = "未找到該筆資料";
            }
            sqlConnection.Close();
            return returns;
        }

        public void UpdateReturns(Returns returns)
        {
            Returns ReverseReturns = new Returns();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseReturns = new SqlCommand("SELECT * FROM returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item");
            sqlCommandReverseReturns.Connection = sqlConnection;
            sqlCommandReverseReturns.Parameters.Add(new SqlParameter("@ReturnsSN", returns.ReturnsSN));
            sqlCommandReverseReturns.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseReturns.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseReturns = new Returns
                    {
                        ReturnsSN = reader.GetString(reader.GetOrdinal("ReturnsSN")),
                        ReturnsDTM = reader.GetDateTime(reader.GetOrdinal("ReturnsDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ReturnsQty = reader.GetInt32(reader.GetOrdinal("ReturnsQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseReturns.ReturnsSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseReturns.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseReturns.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseReturns.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseReturns.ReturnsQty));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT ReturnsQty from returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@ReturnsSN", returns.ReturnsSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseReturns.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseReturns.MoveOutPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //修改退貨單
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE returns SET ReturnsSN = @ReturnsSN, ReturnsDTM = @ReturnsDTM, Item = @Item, PartID = @PartID, ReturnsQty = @ReturnsQty, 
                SupplierID = @SupplierID, MoveOutPlaceID = @MoveOutPlaceID, EmployeeID = @EmployeeID WHERE ReturnsSN = @ReturnsSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsSN", returns.ReturnsSN));
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsDTM", returns.ReturnsDTM));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", returns.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsQty", returns.ReturnsQty));
            sqlCommand.Parameters.Add(new SqlParameter("@SupplierID", returns.SupplierID));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", returns.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", returns.EmployeeID));
            sqlCommand.ExecuteNonQuery();
            

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", returns.ReturnsSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", returns.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", returns.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", returns.MoveOutPlaceID));
            int StockChangeQty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", StockChangeQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", returns.ReturnsQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand SqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT ReturnsQty from returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            SqlCommandStock.Connection = sqlConnection;
            SqlCommandStock.Parameters.Add(new SqlParameter("@ReturnsSN", returns.ReturnsSN));
            SqlCommandStock.Parameters.Add(new SqlParameter("@Item", returns.Item));
            SqlCommandStock.Parameters.Add(new SqlParameter("@PartID", returns.PartID));
            SqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", returns.MoveOutPlaceID));
            SqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public void DeleteReturnsByReturnsSN(string ReturnsSN, int Item)
        {
            Returns ReverseReturns = new Returns();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseReturns = new SqlCommand("SELECT * FROM returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item");
            sqlCommandReverseReturns.Connection = sqlConnection;
            sqlCommandReverseReturns.Parameters.Add(new SqlParameter("@ReturnsSN", ReturnsSN));
            sqlCommandReverseReturns.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseReturns.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseReturns = new Returns
                    {
                        ReturnsSN = reader.GetString(reader.GetOrdinal("ReturnsSN")),
                        ReturnsDTM = reader.GetDateTime(reader.GetOrdinal("ReturnsDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ReturnsQty = reader.GetInt32(reader.GetOrdinal("ReturnsQty")),
                        SupplierID = reader.GetString(reader.GetOrdinal("SupplierID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseReturns.ReturnsSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseReturns.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseReturns.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseReturns.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseReturns.ReturnsQty));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT ReturnsQty from returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@ReturnsSN",ReturnsSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item",Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseReturns.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseReturns.MoveOutPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //刪除退貨單
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM returns WHERE ReturnsSN = @ReturnsSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ReturnsSN", ReturnsSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlCommand.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public List<MaterialRequest> GetMaterialRequests()
        {
            List<MaterialRequest> material_request_list = new List<MaterialRequest>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM materialrequest ORDER BY MaterialRequestSN ASC, Item ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    MaterialRequest material_request = new MaterialRequest
                    {
                        MaterialRequestSN = reader.GetString(reader.GetOrdinal("MaterialRequestSN")),
                        RequestDTM = reader.GetDateTime(reader.GetOrdinal("RequestDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RequestQty = reader.GetInt32(reader.GetOrdinal("RequestQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                    material_request_list.Add(material_request);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return material_request_list;
        }

        public void NewMaterialRequest(MaterialRequest material_request_list)
        {
            //建立領料單
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO materialrequest (MaterialRequestSN, RequestDTM, Item, PartId, RequestQty, MoveOutPlaceID, EmployeeID)
                VALUES (@MaterialRequestSN, @RequestDTM, @Item, @PartID, @RequestQty, @MoveOutPlaceID, @EmployeeID)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@MaterialRequestSN", material_request_list.MaterialRequestSN));
            DateTime date = DateTime.Now;
            sqlCommand.Parameters.Add(new SqlParameter("@RequestDTM", date));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", material_request_list.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", material_request_list.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@RequestQty", material_request_list.RequestQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_request_list.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", material_request_list.EmployeeID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", material_request_list.MaterialRequestSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", material_request_list.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", material_request_list.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_request_list.MoveOutPlaceID));
            int qty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", material_request_list.RequestQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT RequestQty from materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID
                ");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@MaterialRequestSN", material_request_list.MaterialRequestSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", material_request_list.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", material_request_list.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", material_request_list.MoveOutPlaceID));
            sqlCommandStock.ExecuteNonQuery();
            
            sqlConnection.Close();
        }

        public MaterialRequest GetMaterialRequestByMaterialRequestSN(string MaterialRequestSN, int Item)
        {
            MaterialRequest material_request = new MaterialRequest();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@MaterialRequestSN", MaterialRequestSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    material_request = new MaterialRequest
                    {
                        MaterialRequestSN = reader.GetString(reader.GetOrdinal("MaterialRequestSN")),
                        RequestDTM = reader.GetDateTime(reader.GetOrdinal("RequestDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RequestQty = reader.GetInt32(reader.GetOrdinal("RequestQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            else
            {
                material_request.MaterialRequestSN = "未找到該筆資料";
            }
            sqlConnection.Close();
            return material_request;
        }

        public void UpdateMaterialRequest(MaterialRequest material_request)
        {
            MaterialRequest ReverseMaterialRequest = new MaterialRequest();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseMaterialRequest = new SqlCommand("SELECT * FROM materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item");
            sqlCommandReverseMaterialRequest.Connection = sqlConnection;
            sqlCommandReverseMaterialRequest.Parameters.Add(new SqlParameter("@MaterialRequestSN", material_request.MaterialRequestSN));
            sqlCommandReverseMaterialRequest.Parameters.Add(new SqlParameter("@Item", material_request.Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseMaterialRequest.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseMaterialRequest = new MaterialRequest
                    {
                        MaterialRequestSN = reader.GetString(reader.GetOrdinal("MaterialRequestSN")),
                        RequestDTM = reader.GetDateTime(reader.GetOrdinal("RequestDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RequestQty = reader.GetInt32(reader.GetOrdinal("RequestQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            // 沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseMaterialRequest.MaterialRequestSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseMaterialRequest.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialRequest.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseMaterialRequest.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseMaterialRequest.RequestQty));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT RequestQty from materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@MaterialRequestSN", material_request.MaterialRequestSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", material_request.Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialRequest.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialRequest.MoveOutPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //修改領料單
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE materialrequest SET MaterialRequestSN = @MaterialRequestSN, RequestDTM = @RequestDTM, Item = @Item, PartID = @PartID, RequestQty = @RequestQty, 
              MoveOutPlaceID = @MoveOutPlaceID, EmployeeID = @EmployeeID WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@MaterialRequestSN", material_request.MaterialRequestSN));
            sqlCommand.Parameters.Add(new SqlParameter("@RequestDTM", material_request.RequestDTM));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", material_request.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", material_request.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@RequestQty", material_request.RequestQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_request.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", material_request.EmployeeID));            
            sqlCommand.ExecuteNonQuery();

            // 新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", material_request.MaterialRequestSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", material_request.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", material_request.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_request.MoveOutPlaceID));
            int StockChangeQty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", StockChangeQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", material_request.RequestQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT RequestQty from materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@MaterialRequestSN", material_request.MaterialRequestSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", material_request.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialRequest.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialRequest.MoveOutPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public void DeleteMaterialRequestByMaterialRequestSN(string MaterialRequestSN, int Item)
        {
            MaterialRequest ReverseMaterialRequest = new MaterialRequest();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseMaterialRequest = new SqlCommand("SELECT * FROM materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item");
            sqlCommandReverseMaterialRequest.Connection = sqlConnection;
            sqlCommandReverseMaterialRequest.Parameters.Add(new SqlParameter("@MaterialRequestSN", MaterialRequestSN));
            sqlCommandReverseMaterialRequest.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseMaterialRequest.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseMaterialRequest = new MaterialRequest
                    {
                        MaterialRequestSN = reader.GetString(reader.GetOrdinal("MaterialRequestSN")),
                        RequestDTM = reader.GetDateTime(reader.GetOrdinal("RequestDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RequestQty = reader.GetInt32(reader.GetOrdinal("RequestQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();


            // 沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseMaterialRequest.MaterialRequestSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseMaterialRequest.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialRequest.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseMaterialRequest.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseMaterialRequest.RequestQty));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT RequestQty from materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@MaterialRequestSN", MaterialRequestSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialRequest.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialRequest.MoveOutPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //刪除領料單
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM materialrequest WHERE MaterialRequestSN = @MaterialRequestSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@MaterialRequestSN", MaterialRequestSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));            
            sqlCommand.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public List<ReturnToStock> GetReturnToStock()
        {
            List<ReturnToStock> RTS_list = new List<ReturnToStock>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM returntostock ORDER BY RTSSN ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReturnToStock RTSlist = new ReturnToStock
                    {
                        RTSSN = reader.GetString(reader.GetOrdinal("RTSSN")),
                        RTSDTM = reader.GetDateTime(reader.GetOrdinal("RTSDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RTSQty = reader.GetInt32(reader.GetOrdinal("RTSQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                    RTS_list.Add(RTSlist);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return RTS_list;
        }

        public void NewRTS(ReturnToStock RTS_list)
        {
            //建立退料單
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO returntostock (RTSSN, RTSDTM, Item, PartId, RTSQty, MoveInPlaceID, EmployeeID)
                VALUES (@RTSSN, @RTSDTM, @Item, @PartID, @RTSQty, @MoveInPlaceID, @EmployeeID)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@RTSSN", RTS_list.RTSSN));
            DateTime date = DateTime.Now;
            sqlCommand.Parameters.Add(new SqlParameter("@RTSDTM", date));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", RTS_list.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", RTS_list.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@RTSQty", RTS_list.RTSQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveInPlaceID", RTS_list.MoveInPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", RTS_list.EmployeeID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", RTS_list.RTSSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", RTS_list.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", RTS_list.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", RTS_list.MoveInPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", RTS_list.RTSQty));
            int qty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT RTSQty from returntostock WHERE RTSSN = @RTSSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID
                ");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@RTSSN", RTS_list.RTSSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", RTS_list.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", RTS_list.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", RTS_list.MoveInPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public ReturnToStock GetRTSByRTSSN(string RTSSN, int Item)
        {
            ReturnToStock RTS_list = new ReturnToStock();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM returntostock WHERE RTSSN = @RTSSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@RTSSN", RTSSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    RTS_list = new ReturnToStock
                    {
                        RTSSN = reader.GetString(reader.GetOrdinal("RTSSN")),
                        RTSDTM = reader.GetDateTime(reader.GetOrdinal("RTSDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RTSQty = reader.GetInt32(reader.GetOrdinal("RTSQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            else
            {
                RTS_list.RTSSN = "未找到該筆資料";
            }
            sqlConnection.Close();
            return RTS_list;
        }

        public void UpdateRTS(ReturnToStock RTSlist)
        {
            ReturnToStock ReverseRTS_list = new ReturnToStock();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseRTS_list = new SqlCommand("SELECT * FROM returntostock WHERE RTSSN = @RTSSN AND Item = @Item");
            sqlCommandReverseRTS_list.Connection = sqlConnection;
            sqlCommandReverseRTS_list.Parameters.Add(new SqlParameter("@RTSSN", RTSlist.RTSSN));
            sqlCommandReverseRTS_list.Parameters.Add(new SqlParameter("@Item", RTSlist.Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseRTS_list.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseRTS_list = new ReturnToStock
                    {
                        RTSSN = reader.GetString(reader.GetOrdinal("RTSSN")),
                        RTSDTM = reader.GetDateTime(reader.GetOrdinal("RTSDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RTSQty = reader.GetInt32(reader.GetOrdinal("RTSQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseRTS_list.RTSSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseRTS_list.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseRTS_list.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ReverseRTS_list.MoveInPlaceID));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", ReverseRTS_list.RTSQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT RTSQty from returntostock WHERE RTSSN = @RTSSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@RTSSN", RTSlist.RTSSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", RTSlist.Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseRTS_list.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseRTS_list.MoveInPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();


            //修改退料單
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE returntostock SET RTSSN = @RTSSN, RTSDTM = @RTSDTM, Item = @Item, PartID = @PartID, RTSQty = @RTSQty, 
                MoveInPlaceID = @MoveInPlaceID, EmployeeID = @EmployeeID WHERE RTSSN = @RTSSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@RTSSN", RTSlist.RTSSN));
            sqlCommand.Parameters.Add(new SqlParameter("@RTSDTM", RTSlist.RTSDTM));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", RTSlist.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", RTSlist.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@RTSQty", RTSlist.RTSQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveInPlaceID", RTSlist.MoveInPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", RTSlist.EmployeeID));
            
            sqlCommand.ExecuteNonQuery();

            // 新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", RTSlist.RTSSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", RTSlist.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", RTSlist.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", RTSlist.MoveInPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", RTSlist.RTSQty));
            int StockChangeQty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", StockChangeQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT RTSQty from returntostock WHERE RTSSN = @RTSSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@RTSSN", RTSlist.RTSSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", RTSlist.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", RTSlist.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", RTSlist.MoveInPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public void DeleteRTSByRTSSN(string RTSSN, int Item)
        {
            ReturnToStock ReverseRTS_list = new ReturnToStock();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseRTS_list = new SqlCommand("SELECT * FROM returntostock WHERE RTSSN = @RTSSN AND Item = @Item");
            sqlCommandReverseRTS_list.Connection = sqlConnection;
            sqlCommandReverseRTS_list.Parameters.Add(new SqlParameter("@RTSSN",RTSSN));
            sqlCommandReverseRTS_list.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseRTS_list.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseRTS_list = new ReturnToStock
                    {
                        RTSSN = reader.GetString(reader.GetOrdinal("RTSSN")),
                        RTSDTM = reader.GetDateTime(reader.GetOrdinal("RTSDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        RTSQty = reader.GetInt32(reader.GetOrdinal("RTSQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseRTS_list.RTSSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseRTS_list.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseRTS_list.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ReverseRTS_list.MoveInPlaceID));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", ReverseRTS_list.RTSQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT RTSQty from returntostock WHERE RTSSN = @RTSSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@RTSSN", RTSSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseRTS_list.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseRTS_list.MoveInPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //刪除退料單
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM returntostock WHERE RTSSN = @RTSSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@RTSSN", RTSSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlCommand.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public List<MaterialTransfer> GetMaterialTransfers()
        {
            List<MaterialTransfer> material_transfer_list = new List<MaterialTransfer>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM materialtransfer ORDER BY TransferSN ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    MaterialTransfer material_transfer = new MaterialTransfer
                    {
                        TransferSN = reader.GetString(reader.GetOrdinal("TransferSN")),
                        TransferDTM = reader.GetDateTime(reader.GetOrdinal("TransferDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        TransferQty = reader.GetInt32(reader.GetOrdinal("TransferQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                    material_transfer_list.Add(material_transfer);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return material_transfer_list;
        }

        public void NewMaterialTransfer(MaterialTransfer material_transfer_list)
        {
            //建立調料單
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO materialtransfer (TransferSN, TransferDTM, Item, PartId, TransferQty, MoveInPlaceID, MoveOutPlaceID, EmployeeID)
                VALUES (@TransferSN, @TransferDTM, @Item, @PartID, @TransferQty, @MoveInPlaceID, @MoveOutPlaceID, @EmployeeID)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@TransferSN", material_transfer_list.TransferSN));
            DateTime date = DateTime.Now;
            sqlCommand.Parameters.Add(new SqlParameter("@TransferDTM", date));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", material_transfer_list.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", material_transfer_list.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@TransferQty", material_transfer_list.TransferQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveInPlaceID", material_transfer_list.MoveInPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_transfer_list.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", material_transfer_list.EmployeeID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();

            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", material_transfer_list.TransferSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", material_transfer_list.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", material_transfer_list.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", material_transfer_list.MoveInPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_transfer_list.MoveOutPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", material_transfer_list.TransferQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", material_transfer_list.TransferQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新移出倉別庫存
            SqlCommand sqlCommandStockOut = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID
                ");
            sqlCommandStockOut.Connection = sqlConnection;
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@TransferSN", material_transfer_list.TransferSN));
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@Item", material_transfer_list.Item));
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@PartID", material_transfer_list.PartID));
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@PlaceID", material_transfer_list.MoveOutPlaceID));
            sqlCommandStockOut.ExecuteNonQuery();

            //更新移入倉別庫存
            SqlCommand sqlCommandStockIn = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID
                ");
            sqlCommandStockIn.Connection = sqlConnection;
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@TransferSN", material_transfer_list.TransferSN));
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@Item", material_transfer_list.Item));
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@PartID", material_transfer_list.PartID));
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@PlaceID", material_transfer_list.MoveInPlaceID));
            sqlCommandStockIn.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public MaterialTransfer GetMaterialTransferByTransferSN(string TransferSN, int Item)
        {
            MaterialTransfer material_transfer = new MaterialTransfer();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@TransferSN", TransferSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    material_transfer = new MaterialTransfer
                    {
                        TransferSN = reader.GetString(reader.GetOrdinal("TransferSN")),
                        TransferDTM = reader.GetDateTime(reader.GetOrdinal("TransferDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        TransferQty = reader.GetInt32(reader.GetOrdinal("TransferQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            else
            {
                material_transfer.TransferSN = "未找到該筆資料";
            }
            sqlConnection.Close();
            return material_transfer;
        }

        public void UpdateMaterialTransfer(MaterialTransfer material_transfer)
        {
            MaterialTransfer ReverseMaterialTransfer = new MaterialTransfer();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseMaterialTransfer = new SqlCommand("SELECT * FROM materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item");
            sqlCommandReverseMaterialTransfer.Connection = sqlConnection;
            sqlCommandReverseMaterialTransfer.Parameters.Add(new SqlParameter("@TransferSN", material_transfer.TransferSN));
            sqlCommandReverseMaterialTransfer.Parameters.Add(new SqlParameter("@Item", material_transfer.Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseMaterialTransfer.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseMaterialTransfer = new MaterialTransfer
                    {
                        TransferSN = reader.GetString(reader.GetOrdinal("TransferSN")),
                        TransferDTM = reader.GetDateTime(reader.GetOrdinal("TransferDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        TransferQty = reader.GetInt32(reader.GetOrdinal("TransferQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseMaterialTransfer.TransferSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseMaterialTransfer.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialTransfer.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseMaterialTransfer.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ReverseMaterialTransfer.MoveInPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseMaterialTransfer.TransferQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", ReverseMaterialTransfer.TransferQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正移入倉別庫存
            SqlCommand sqlCommandReverseStockIn = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStockIn.Connection = sqlConnection;
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@TransferSN", ReverseMaterialTransfer.TransferSN));
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@Item", ReverseMaterialTransfer.Item));
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialTransfer.PartID));
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialTransfer.MoveInPlaceID));
            sqlCommandReverseStockIn.ExecuteNonQuery();

            //沖正移出倉別庫存
            SqlCommand sqlCommandReverseStockOut = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStockOut.Connection = sqlConnection;
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@TransferSN", ReverseMaterialTransfer.TransferSN));
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@Item", ReverseMaterialTransfer.Item));
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialTransfer.PartID));
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialTransfer.MoveOutPlaceID));
            sqlCommandReverseStockOut.ExecuteNonQuery();

            //修改調料單
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE materialtransfer SET TransferSN = @TransferSN, TransferDTM = @TransferDTM, Item = @Item, PartID = @PartID, TransferQty = @TransferQty, 
                MoveInPlaceID = @MoveInPlaceID, MoveOutPlaceID = @MoveOutPlaceID, EmployeeID = @EmployeeID WHERE TransferSN = @TransferSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@TransferSN", material_transfer.TransferSN));
            sqlCommand.Parameters.Add(new SqlParameter("@TransferDTM", material_transfer.TransferDTM));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", material_transfer.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", material_transfer.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@TransferQty", material_transfer.TransferQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveInPlaceID", material_transfer.MoveInPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_transfer.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", material_transfer.EmployeeID));
            sqlCommand.ExecuteNonQuery();

            // 新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", material_transfer.TransferSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", material_transfer.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", material_transfer.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", material_transfer.MoveInPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", material_transfer.MoveOutPlaceID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", material_transfer.TransferQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", material_transfer.TransferQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新移出倉別庫存
            SqlCommand sqlCommandStockOut = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandStockOut.Connection = sqlConnection;
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@TransferSN", material_transfer.TransferSN));
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@Item", material_transfer.Item));
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@PartID", material_transfer.PartID));
            sqlCommandStockOut.Parameters.Add(new SqlParameter("@PlaceID", material_transfer.MoveOutPlaceID));
            sqlCommandStockOut.ExecuteNonQuery();

            //更新移入倉別庫存
            SqlCommand sqlCommandStockIn = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandStockIn.Connection = sqlConnection;
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@TransferSN", material_transfer.TransferSN));
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@Item", material_transfer.Item));
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@PartID", material_transfer.PartID));
            sqlCommandStockIn.Parameters.Add(new SqlParameter("@PlaceID", material_transfer.MoveInPlaceID));
            sqlCommandStockIn.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public void DeleteMaterialTransferByTransferSN(string TransferSN, int Item)
        {
            MaterialTransfer ReverseMaterialTransfer = new MaterialTransfer();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseMaterialTransfer = new SqlCommand("SELECT * FROM materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item");
            sqlCommandReverseMaterialTransfer.Connection = sqlConnection;
            sqlCommandReverseMaterialTransfer.Parameters.Add(new SqlParameter("@TransferSN", TransferSN));
            sqlCommandReverseMaterialTransfer.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseMaterialTransfer.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseMaterialTransfer = new MaterialTransfer
                    {
                        TransferSN = reader.GetString(reader.GetOrdinal("TransferSN")),
                        TransferDTM = reader.GetDateTime(reader.GetOrdinal("TransferDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        TransferQty = reader.GetInt32(reader.GetOrdinal("TransferQty")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseMaterialTransfer.TransferSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseMaterialTransfer.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialTransfer.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseMaterialTransfer.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ReverseMaterialTransfer.MoveInPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseMaterialTransfer.TransferQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", ReverseMaterialTransfer.TransferQty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正移入倉別庫存
            SqlCommand sqlCommandReverseStockIn = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStockIn.Connection = sqlConnection;
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@TransferSN", ReverseMaterialTransfer.TransferSN));
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@Item", ReverseMaterialTransfer.Item));
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialTransfer.PartID));
            sqlCommandReverseStockIn.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialTransfer.MoveInPlaceID));
            sqlCommandReverseStockIn.ExecuteNonQuery();

            //沖正移出倉別庫存
            SqlCommand sqlCommandReverseStockOut = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT TransferQty from materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStockOut.Connection = sqlConnection;
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@TransferSN", ReverseMaterialTransfer.TransferSN));
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@Item", ReverseMaterialTransfer.Item));
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@PartID", ReverseMaterialTransfer.PartID));
            sqlCommandReverseStockOut.Parameters.Add(new SqlParameter("@PlaceID", ReverseMaterialTransfer.MoveOutPlaceID));
            sqlCommandReverseStockOut.ExecuteNonQuery();

            //刪除調料單
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM materialtransfer WHERE TransferSN = @TransferSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@TransferSN", TransferSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }

        public List<Scrap> GetScraps()
        {
            List<Scrap> scrap_list = new List<Scrap>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM scrap ORDER BY ScrapSN ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    Scrap scraps = new Scrap
                    {
                        ScrapSN = reader.GetString(reader.GetOrdinal("ScrapSN")),
                        ScrapDTM = reader.GetDateTime(reader.GetOrdinal("ScrapDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ScrapQty = reader.GetInt32(reader.GetOrdinal("ScrapQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                    scrap_list.Add(scraps);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return scrap_list;
        }

        public void NewScrap(Scrap scrap_list)
        {
            //新增報廢單
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"INSERT INTO scrap (ScrapSN, ScrapDTM, Item, PartId, ScrapQty, MoveOutPlaceID, EmployeeID)
                VALUES (@ScrapSN, @ScrapDTM, @Item, @PartID, @ScrapQty, @MoveOutPlaceID, @EmployeeID)");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapSN", scrap_list.ScrapSN));
            DateTime date = DateTime.Now;
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapDTM", date));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", scrap_list.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", scrap_list.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapQty", scrap_list.ScrapQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", scrap_list.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", scrap_list.EmployeeID));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", scrap_list.ScrapSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", scrap_list.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", scrap_list.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", scrap_list.MoveOutPlaceID));
            int qty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", qty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", scrap_list.ScrapQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT ScrapQty from scrap WHERE ScrapSN = @ScrapSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID
                ");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@ScrapSN", scrap_list.ScrapSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", scrap_list.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", scrap_list.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", scrap_list.MoveOutPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public Scrap GetScrapByScrapSN(string ScrapSN, int Item)
        {
            Scrap scraps = new Scrap();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM scrap WHERE ScrapSN = @ScrapSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapSN", ScrapSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    scraps = new Scrap
                    {
                        ScrapSN = reader.GetString(reader.GetOrdinal("ScrapSN")),
                        ScrapDTM = reader.GetDateTime(reader.GetOrdinal("ScrapDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ScrapQty = reader.GetInt32(reader.GetOrdinal("ScrapQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            else
            {
                scraps.ScrapSN = "未找到該筆資料";
            }
            sqlConnection.Close();
            return scraps;
        }

        public void UpdateScrap(Scrap scraps)
        {
            Scrap ReverseScraps = new Scrap();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseScraps = new SqlCommand("SELECT * FROM scrap WHERE ScrapSN = @ScrapSN AND Item = @Item");
            sqlCommandReverseScraps.Connection = sqlConnection;
            sqlCommandReverseScraps.Parameters.Add(new SqlParameter("@ScrapSN", scraps.ScrapSN));
            sqlCommandReverseScraps.Parameters.Add(new SqlParameter("@Item", scraps.Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseScraps.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseScraps = new Scrap
                    {
                        ScrapSN = reader.GetString(reader.GetOrdinal("ScrapSN")),
                        ScrapDTM = reader.GetDateTime(reader.GetOrdinal("ScrapDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ScrapQty = reader.GetInt32(reader.GetOrdinal("ScrapQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseScraps.ScrapSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseScraps.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseScraps.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseScraps.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseScraps.ScrapQty));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT ScrapQty from scrap WHERE ScrapSN = @ScrapSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@ScrapSN", ReverseScraps.ScrapSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", ReverseScraps.Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseScraps.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseScraps.MoveOutPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();


            //修改報廢單
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE scrap SET ScrapSN = @ScrapSN, ScrapDTM = @ScrapDTM, Item = @Item, PartID = @PartID, ScrapQty = @ScrapQty, 
              MoveOutPlaceID = @MoveOutPlaceID, EmployeeID = @EmployeeID WHERE ScrapSN = @ScrapSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapSN", scraps.ScrapSN));
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapDTM", scraps.ScrapDTM));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", scraps.Item));
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", scraps.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapQty", scraps.ScrapQty));
            sqlCommand.Parameters.Add(new SqlParameter("@MoveOutPlaceID", scraps.MoveOutPlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@EmployeeID", scraps.EmployeeID));
            sqlCommand.ExecuteNonQuery();

            //新增庫存異動明細
            SqlCommand sqlCommandStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandStockChange.Connection = sqlConnection;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", scraps.ScrapSN));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@Item", scraps.Item));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@PartID", scraps.PartID));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ""));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", scraps.MoveOutPlaceID));
            int StockChangeQty = 0;
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveInQty", StockChangeQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", scraps.ScrapQty));
            sqlCommandStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "正常"));
            sqlCommandStockChange.ExecuteNonQuery();

            //更新庫存
            SqlCommand sqlCommandStock = new SqlCommand(
                @"UPDATE stock SET Qty -= (SELECT ScrapQty from scrap WHERE ScrapSN = @ScrapSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandStock.Connection = sqlConnection;
            sqlCommandStock.Parameters.Add(new SqlParameter("@ScrapSN", scraps.ScrapSN));
            sqlCommandStock.Parameters.Add(new SqlParameter("@Item", scraps.Item));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PartID", scraps.PartID));
            sqlCommandStock.Parameters.Add(new SqlParameter("@PlaceID", scraps.MoveOutPlaceID));
            sqlCommandStock.ExecuteNonQuery();

            sqlConnection.Close();
        }

        public void DeleteScrapByScrapSN(string ScrapSN, int Item)
        {
            Scrap ReverseScraps = new Scrap();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommandReverseScraps = new SqlCommand("SELECT * FROM scrap WHERE ScrapSN = @ScrapSN AND Item = @Item");
            sqlCommandReverseScraps.Connection = sqlConnection;
            sqlCommandReverseScraps.Parameters.Add(new SqlParameter("@ScrapSN", ScrapSN));
            sqlCommandReverseScraps.Parameters.Add(new SqlParameter("@Item", Item));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommandReverseScraps.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ReverseScraps = new Scrap
                    {
                        ScrapSN = reader.GetString(reader.GetOrdinal("ScrapSN")),
                        ScrapDTM = reader.GetDateTime(reader.GetOrdinal("ScrapDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        ScrapQty = reader.GetInt32(reader.GetOrdinal("ScrapQty")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        EmployeeID = reader.GetString(reader.GetOrdinal("EmployeeID"))
                    };
                }
            }
            sqlConnection.Close();

            //沖正庫存異動明細
            SqlCommand sqlCommandReverseStockChange = new SqlCommand(
                @"INSERT INTO stockchange (StockChangeSN, StockChangeDTM, Item, PartID, MoveInPlaceID, MoveOutPlaceID, MoveInQty, MoveOutQty, ChangeStatusType)
                VALUES (@StockChangeSN, @StockChangeDTM, @Item, @PartID, @MoveInPlaceID, @MoveOutPlaceID, @MoveInQty, @MoveOutQty, @ChangeStatusType)");
            sqlCommandReverseStockChange.Connection = sqlConnection;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeSN", ReverseScraps.ScrapSN));
            DateTime date = DateTime.Now;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@StockChangeDTM", date));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@Item", ReverseScraps.Item));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@PartID", ReverseScraps.PartID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInPlaceID", ReverseScraps.MoveOutPlaceID));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutPlaceID", ""));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveInQty", ReverseScraps.ScrapQty));
            int qty = 0;
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@MoveOutQty", qty));
            sqlCommandReverseStockChange.Parameters.Add(new SqlParameter("@ChangeStatusType", "沖正"));
            sqlConnection.Open();
            sqlCommandReverseStockChange.ExecuteNonQuery();

            //沖正庫存
            SqlCommand sqlCommandReverseStock = new SqlCommand(
                @"UPDATE stock SET Qty += (SELECT ScrapQty from scrap WHERE ScrapSN = @ScrapSN AND Item = @Item)
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommandReverseStock.Connection = sqlConnection;
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@ScrapSN", ScrapSN));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@Item", Item));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PartID", ReverseScraps.PartID));
            sqlCommandReverseStock.Parameters.Add(new SqlParameter("@PlaceID", ReverseScraps.MoveOutPlaceID));
            sqlCommandReverseStock.ExecuteNonQuery();

            //刪除報廢單
            SqlCommand sqlCommand = new SqlCommand("DELETE FROM scrap WHERE ScrapSN = @ScrapSN AND Item = @Item");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@ScrapSN", ScrapSN));
            sqlCommand.Parameters.Add(new SqlParameter("@Item", Item));
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }

        public List<StockChange> GetStockChange()
        {
            List<StockChange> stock_change_list = new List<StockChange>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM stockchange ORDER BY StockChangeSN ASC, StockChangeDTM ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    StockChange stock_change = new StockChange
                    {
                        StockChangeSN = reader.GetString(reader.GetOrdinal("StockChangeSN")),
                        StockChangeDTM = reader.GetDateTime(reader.GetOrdinal("StockChangeDTM")),
                        Item = reader.GetInt32(reader.GetOrdinal("Item")),
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        MoveInPlaceID = reader.GetString(reader.GetOrdinal("MoveInPlaceID")),
                        MoveOutPlaceID = reader.GetString(reader.GetOrdinal("MoveOutPlaceID")),
                        MoveInQty = reader.GetInt32(reader.GetOrdinal("MoveInQty")),
                        MoveOutQty = reader.GetInt32(reader.GetOrdinal("MoveOutQty")),
                        ChangeStatusType = reader.GetString(reader.GetOrdinal("ChangeStatusType"))
                    };
                    stock_change_list.Add(stock_change);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return stock_change_list;
        }

        public List<Stock> GetStock()
        {
            List<Stock> stock_list = new List<Stock>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT A.PartID, B.PartName, A.PlaceID, A.Qty FROM Stock AS A JOIN Part AS B ON (A.PartID = B.PartID) ORDER BY PlaceID ASC, PartID ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    Stock stock = new Stock
                    {
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        PartName = reader.GetString(reader.GetOrdinal("PartName")),
                        PlaceID = reader.GetString(reader.GetOrdinal("PlaceID")),
                        Qty = reader.GetInt32(reader.GetOrdinal("Qty"))
                    };
                    stock_list.Add(stock);
                }
            }
            else
            {
                Console.WriteLine("查無資料");
            }
            sqlConnection.Close();
            return stock_list;
        }

        public List<PhysicalCountList> GetPhysicalCountList()
        {
            List<PhysicalCountList> count_list = new List<PhysicalCountList>();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT A.PartID, C.PartName, A.PlaceID, B.Qty, A.CountedQty, A.Difference FROM PhysicalCount AS A JOIN Stock AS B ON (A.PartID = B.PartID AND A.PlaceID = B.PlaceID) JOIN Part AS C ON (A.PartID = C.PartID) ORDER BY A.PlaceID ASC, A.PartID ASC");
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    PhysicalCountList CountList = new PhysicalCountList
                    {
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        PartName = reader.GetString(reader.GetOrdinal("PartName")),
                        PlaceID = reader.GetString(reader.GetOrdinal("PlaceID")),
                        SystemQty = reader.GetInt32(reader.GetOrdinal("Qty")),
                        CountedQty = reader.GetInt32(reader.GetOrdinal("CountedQty")),
                        Difference = reader.GetInt32(reader.GetOrdinal("Difference")),
                    };
                    
                    CountList.Difference = CountList.CountedQty - CountList.SystemQty;
                    count_list.Add(CountList);
                   
                }
            }
            //reader.Close();
            //SqlCommand sqlCommandSystemQty = new SqlCommand(
            //         @"UPDATE PhysicalCount 
            //            SET SystemQty = (SELECT Qty FROM stock WHERE PartID = @PartID AND PlaceID = @PlaceID),
            //            CountedQty = (SELECT Qty FROM stock WHERE PartID = @PartID AND PlaceID = @PlaceID),
            //            Difference = CountedQty - SystemQty
            //            WHERE PartID = @PartID AND PlaceID = @PlaceID");

            //sqlCommandSystemQty.Connection = sqlConnection;
            //sqlCommandSystemQty.Parameters.Add(new SqlParameter("@PartID", count_list.PartID));
            //sqlCommandSystemQty.Parameters.Add(new SqlParameter("@PlaceID", count_list.PlaceID));
            //sqlCommandSystemQty.ExecuteNonQuery();
            sqlConnection.Close();
            return count_list;
        }

        public PhysicalCountList GetPhysicalCountListByPartID(string PartID, string PlaceID)
        {
            PhysicalCountList CountList = new PhysicalCountList();
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand("SELECT A.PartID, C.PartName, A.PlaceID, B.Qty, A.CountedQty, A.Difference FROM PhysicalCount AS A JOIN Stock AS B ON (A.PartID = B.PartID AND A.PlaceID = B.PlaceID) JOIN Part AS C ON (A.PartID = C.PartID) WHERE A.PartID = @PartID AND A.PlaceID = @PlaceID");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceID", PlaceID));
            sqlConnection.Open();

            SqlDataReader reader = sqlCommand.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    CountList = new PhysicalCountList
                    {
                        PartID = reader.GetString(reader.GetOrdinal("PartID")),
                        PartName = reader.GetString(reader.GetOrdinal("PartName")),
                        PlaceID = reader.GetString(reader.GetOrdinal("PlaceID")),
                        SystemQty = reader.GetInt32(reader.GetOrdinal("Qty")),
                        CountedQty = reader.GetInt32(reader.GetOrdinal("CountedQty")),
                        Difference = reader.GetInt32(reader.GetOrdinal("CountedQty"))
                    };
                    CountList.Difference = CountList.CountedQty - CountList.SystemQty;
                }
            }
            else
            {
                CountList.PartID = "未找到該筆資料";
            }
            sqlConnection.Close();
            return CountList;
        }

            public void UpdatePhysicalCountList(PhysicalCountList CountList)
        {
            //更新盤點清冊
            SqlConnection sqlConnection = new SqlConnection(ConnStr);
            SqlCommand sqlCommand = new SqlCommand(
                @"UPDATE PhysicalCount SET PartID = @PartID, PlaceID = @PlaceID, SystemQty = @SystemQty, CountedQty = @CountedQty, Difference = @Difference
                WHERE PartID = @PartID AND PlaceID = @PlaceID");
            sqlCommand.Connection = sqlConnection;
            sqlCommand.Parameters.Add(new SqlParameter("@PartID", CountList.PartID));
            sqlCommand.Parameters.Add(new SqlParameter("@PlaceID", CountList.PlaceID));
            sqlCommand.Parameters.Add(new SqlParameter("@SystemQty", CountList.SystemQty));
            sqlCommand.Parameters.Add(new SqlParameter("@CountedQty", CountList.CountedQty));
            int Difference = CountList.CountedQty - CountList.SystemQty;
            sqlCommand.Parameters.Add(new SqlParameter("@Difference", Difference));
            sqlConnection.Open();
            sqlCommand.ExecuteNonQuery();
            sqlConnection.Close();
        }

    }



    public class Place
    {
        public string PlaceID { get; set; }
        public string PlaceName { get; set; }
    }
    
    public class GoodsReceipt
    {
        public string GRSN { get; set; }
        public DateTime ReceiptDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public int GRQty { get; set; }
        public string SupplierID { get; set; }
        public string MoveInPlaceID { get; set; }
        public string EmployeeID { get; set; }
    }

    public class Returns
    {
        public string ReturnsSN { get; set; }
        public DateTime ReturnsDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public int ReturnsQty { get; set; }
        public string SupplierID { get; set; }
        public string MoveOutPlaceID { get; set; }
        public string EmployeeID { get; set; }
    }

    public class MaterialRequest
    {
        public string MaterialRequestSN { get; set; }
        public DateTime RequestDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public int RequestQty { get; set; }
        public string MoveOutPlaceID { get; set; }
        public string EmployeeID { get; set; }
    }

    public class ReturnToStock
    {
        public string RTSSN { get; set; }
        public DateTime RTSDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public int RTSQty { get; set; }
        public string MoveInPlaceID { get; set; }
        public string EmployeeID { get; set; }
    }

    public class MaterialTransfer
    {
        public string TransferSN { get; set; }
        public DateTime TransferDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public int TransferQty { get; set; }
        public string MoveInPlaceID { get; set; }
        public string MoveOutPlaceID { get; set; }
        public string EmployeeID { get; set; }
    }

    public class Scrap
    {
        public string ScrapSN { get; set; }
        public DateTime ScrapDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public int ScrapQty { get; set; }
        public string MoveOutPlaceID { get; set; }
        public string EmployeeID { get; set; }
    }

    public class StockChange
    {
        public string StockChangeSN { get; set; }
        public DateTime StockChangeDTM { get; set; }
        public int Item { get; set; }
        public string PartID { get; set; }
        public string MoveInPlaceID { get; set; }
        public string MoveOutPlaceID { get; set; }
        public int MoveInQty { get; set; }
        public int MoveOutQty { get; set; }
        public string ChangeStatusType { get; set; }
    }

    public class Stock
    {
        public string PartID { get; set; }
        public string PartName { get; set; }
        public string PlaceID { get; set; }
        public int Qty { get; set; }
    }

    public class PhysicalCountList
    {
        public string PartID { get; set; }
        public string PartName { get; set; }
        public string PlaceID { get; set; }
        public int SystemQty { get; set; }
        public int CountedQty { get; set; }
        public int Difference { get; set; }
    }
}