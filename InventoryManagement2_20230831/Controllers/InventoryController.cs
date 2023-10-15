using InventoryManagement3.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace InventoryManagement3.Controllers
{    
    public class InventoryController : Controller
    {
        // GET: Inventory
        public ActionResult Place()
        {
            DBmanager dbmanager = new DBmanager();
            List<Place> place_list = dbmanager.GetPlaces();
            ViewBag.place_list = place_list;
            return View();
        }

        public ActionResult CreatePlace()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreatePlace(Place place_list)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewPlace(place_list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("Place");
        }

        public ActionResult EditPlace(string PlaceID)
        {
            DBmanager dBmanager = new DBmanager();
            Place place = dBmanager.GetPlaceByPlaceID(PlaceID);
            return View(place);
        }

        [HttpPost]
        public ActionResult EditPlace(Place place)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdatePlace(place);
            return RedirectToAction("Place");
        }

        public ActionResult DeletePlace(string PlaceID)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeletePlaceByPlaceID(PlaceID);
            return RedirectToAction("Place");
        }

        public ActionResult GoodsReceipt()
        {
            DBmanager dbmanager = new DBmanager();
            List<GoodsReceipt> goods_receipt = dbmanager.GetGoodsReceipts();
            ViewBag.goods_receipt = goods_receipt;
            return View();
        }

        public ActionResult CreateGoodsReceipt()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreateGoodsReceipt(GoodsReceipt goods_receipt)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewGoodsReceipt(goods_receipt);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("GoodsReceipt");
        }



        public ActionResult EditGoodsReceipt(string GRSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            GoodsReceipt goodsreceipt = dBmanager.GetGoodsReceiptByGRSN(GRSN, Item);
            return View(goodsreceipt);
        }

        [HttpPost]
        public ActionResult EditGoodsReceipt(GoodsReceipt goodsreceipt)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdateGoodsReceipt(goodsreceipt);
            return RedirectToAction("GoodsReceipt");
        }

        public ActionResult DeleteGoodsReceipt(string GRSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeleteGoodsReceiptByGRSN(GRSN, Item);
            return RedirectToAction("GoodsReceipt");
        }

        public ActionResult Returns()
        {
            DBmanager dBmanager = new DBmanager();
            List<Returns> returns_list = dBmanager.GetReturns();
            ViewBag.returns_list = returns_list;
            return View();
        }

        public ActionResult CreateReturns()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreateReturns(Returns returns_list)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewReturns(returns_list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("Returns");
        }

        public ActionResult EditReturns(string ReturnsSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            Returns returns = dBmanager.GetReturnsByReturnsSN(ReturnsSN, Item);
            return View(returns);
        }

        [HttpPost]
        public ActionResult EditReturns(Returns returns)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdateReturns(returns);
            return RedirectToAction("Returns");

        }

        public ActionResult DeleteReturns(string ReturnsSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeleteReturnsByReturnsSN(ReturnsSN, Item);
            return RedirectToAction("Returns");
        }

        public ActionResult MaterialRequest()
        {
            DBmanager dBmanager = new DBmanager();
            List<MaterialRequest> material_request_list = dBmanager.GetMaterialRequests();
            ViewBag.material_request_list = material_request_list;
            return View();
        }

        public ActionResult CreateMaterialRequest()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreateMaterialRequest(MaterialRequest material_request_list)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewMaterialRequest(material_request_list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("MaterialRequest");
        }

        public ActionResult EditMaterialRequest(string MaterialRequestSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            MaterialRequest material_request = dBmanager.GetMaterialRequestByMaterialRequestSN(MaterialRequestSN, Item);
            return View(material_request);
        }

        [HttpPost]
        public ActionResult EditMaterialRequest(MaterialRequest material_request_list)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdateMaterialRequest(material_request_list);
            return RedirectToAction("MaterialRequest");

        }

        public ActionResult DeleteMaterialRequest(string MaterialRequestSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeleteMaterialRequestByMaterialRequestSN(MaterialRequestSN, Item);
            return RedirectToAction("MaterialRequest");
        }

        public ActionResult RTS()
        {
            DBmanager dBmanager = new DBmanager();
            List<ReturnToStock> RTS_list = dBmanager.GetReturnToStock();
            ViewBag.RTS_list = RTS_list;
            return View();
        }

        public ActionResult CreateRTS()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreateRTS(ReturnToStock RTS_list)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewRTS(RTS_list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("RTS");
        }

        public ActionResult EditRTS(string RTSSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            ReturnToStock RTSlist = dBmanager.GetRTSByRTSSN(RTSSN, Item);
            return View(RTSlist);
        }

        [HttpPost]
        public ActionResult EditRTS(ReturnToStock RTS_list)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdateRTS(RTS_list);
            return RedirectToAction("RTS");

        }

        public ActionResult DeleteRTS(string RTSSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeleteRTSByRTSSN(RTSSN, Item);
            return RedirectToAction("RTS");
        }

        public ActionResult MaterialTransfer()
        {
            DBmanager dBmanager = new DBmanager();
            List<MaterialTransfer> material_transfer_list = dBmanager.GetMaterialTransfers();
            ViewBag.material_transfer_list = material_transfer_list;
            return View();
        }

        public ActionResult CreateMaterialTransfer()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreateMaterialTransfer(MaterialTransfer material_transfer_list)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewMaterialTransfer(material_transfer_list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("MaterialTransfer");
        }

        public ActionResult EditMaterialTransfer(string TransferSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            MaterialTransfer material_transfer = dBmanager.GetMaterialTransferByTransferSN(TransferSN, Item);
            return View(material_transfer);
        }

        [HttpPost]
        public ActionResult EditMaterialTransfer(MaterialTransfer material_transfer_list)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdateMaterialTransfer(material_transfer_list);
            return RedirectToAction("MaterialTransfer");

        }

        public ActionResult DeleteMaterialTransfer(string TransferSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeleteMaterialTransferByTransferSN(TransferSN, Item);
            return RedirectToAction("MaterialTransfer");
        }

        public ActionResult Scrap()
        {
            DBmanager dBmanager = new DBmanager();
            List<Scrap> scrap_list = dBmanager.GetScraps();
            ViewBag.scrap_list = scrap_list;
            return View();
        }

        public ActionResult CreateScrap()
        {
            return View();
        }

        [HttpPost]
        public ActionResult CreateScrap(Scrap scrap_list)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.NewScrap(scrap_list);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("Scrap");
        }

        public ActionResult EditScrap(string ScrapSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            Scrap scraps = dBmanager.GetScrapByScrapSN(ScrapSN, Item);
            return View(scraps);
        }

        [HttpPost]
        public ActionResult EditScrap(Scrap scrap_list)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.UpdateScrap(scrap_list);
            return RedirectToAction("Scrap");
        }

        public ActionResult DeleteScrap(string ScrapSN, int Item)
        {
            DBmanager dBmanager = new DBmanager();
            dBmanager.DeleteScrapByScrapSN(ScrapSN, Item);
            return RedirectToAction("Scrap");
        }

        public ActionResult Stock()
        {
            DBmanager dBmanager = new DBmanager();
            List<Stock> stock_list = dBmanager.GetStock();
            ViewBag.stock_list = stock_list;
            return View();
        }

        public ActionResult StockChange()
        {
            DBmanager dBmanager = new DBmanager();
            List<StockChange> stock_change_list = dBmanager.GetStockChange();
            ViewBag.stock_change_list = stock_change_list;
            return View();
        }

        public ActionResult PhysicalCountList()
        {
            DBmanager dBmanager = new DBmanager();
            List<PhysicalCountList> count_list = dBmanager.GetPhysicalCountList();
            ViewBag.count_list = count_list;
            return View();
        }

        public ActionResult EditPhysicalCountList(string PartID, string PlaceID)
        {
            DBmanager dBmanager = new DBmanager();
            PhysicalCountList CountList = dBmanager.GetPhysicalCountListByPartID(PartID, PlaceID);
            return View(CountList);
        }

        [HttpPost]
        public ActionResult EditPhysicalCountList(PhysicalCountList CountList)
        {
            DBmanager dbmanager = new DBmanager();
            try
            {
                dbmanager.UpdatePhysicalCountList(CountList);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return RedirectToAction("PhysicalCountList");
        }


    }
}