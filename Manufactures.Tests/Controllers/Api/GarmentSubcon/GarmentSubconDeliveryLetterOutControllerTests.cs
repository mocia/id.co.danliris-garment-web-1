﻿using Barebone.Tests;
using Manufactures.Controllers.Api.GarmentSubcon;
using Manufactures.Domain.GarmentSubcon.SubconDeliveryLetterOuts;
using Manufactures.Domain.GarmentSubcon.SubconDeliveryLetterOuts.Commands;
using Manufactures.Domain.GarmentSubcon.SubconDeliveryLetterOuts.ReadModels;
using Manufactures.Domain.GarmentSubcon.SubconDeliveryLetterOuts.Repositories;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Security.Claims;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace Manufactures.Tests.Controllers.Api.GarmentSubcon
{
    public class GarmentSubconDeliveryLetterOutControllerTests : BaseControllerUnitTest
    {
        private Mock<IGarmentSubconDeliveryLetterOutRepository> _mockGarmentSubconDeliveryLetterOutRepository;
        private Mock<IGarmentSubconDeliveryLetterOutItemRepository> _mockGarmentSubconDeliveryLetterOutItemRepository;

        public GarmentSubconDeliveryLetterOutControllerTests() : base()
        {
            _mockGarmentSubconDeliveryLetterOutRepository = CreateMock<IGarmentSubconDeliveryLetterOutRepository>();
            _mockGarmentSubconDeliveryLetterOutItemRepository = CreateMock<IGarmentSubconDeliveryLetterOutItemRepository>();

            _MockStorage.SetupStorage(_mockGarmentSubconDeliveryLetterOutRepository);
            _MockStorage.SetupStorage(_mockGarmentSubconDeliveryLetterOutItemRepository);

        }

        private GarmentSubconDeliveryLetterOutController CreateGarmentSubconDeliveryLetterOutController()
        {
            var user = new Mock<ClaimsPrincipal>();
            var claims = new Claim[]
            {
                new Claim("username", "unittestusername")
            };
            user.Setup(u => u.Claims).Returns(claims);
            GarmentSubconDeliveryLetterOutController controller = (GarmentSubconDeliveryLetterOutController)Activator.CreateInstance(typeof(GarmentSubconDeliveryLetterOutController), _MockServiceProvider.Object);
            controller.ControllerContext = new ControllerContext()
            {
                HttpContext = new DefaultHttpContext()
                {
                    User = user.Object
                }
            };
            controller.ControllerContext.HttpContext.Request.Headers["Authorization"] = "Bearer unittesttoken";
            controller.ControllerContext.HttpContext.Request.Headers["x-timezone-offset"] = "7";
            controller.ControllerContext.HttpContext.Request.Path = new PathString("/v1/unit-test");
            return controller;
        }

        [Fact]
        public async Task Get_StateUnderTest_ExpectedBehavior()
        {
            // Arrange
            var unitUnderTest = CreateGarmentSubconDeliveryLetterOutController();

            _mockGarmentSubconDeliveryLetterOutRepository
                .Setup(s => s.Read(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>()))
                .Returns(new List<GarmentSubconDeliveryLetterOutReadModel>().AsQueryable());

            Guid SubconDeliveryLetterOutGuid = Guid.NewGuid();
            _mockGarmentSubconDeliveryLetterOutRepository
                .Setup(s => s.Find(It.IsAny<IQueryable<GarmentSubconDeliveryLetterOutReadModel>>()))
                .Returns(new List<GarmentSubconDeliveryLetterOut>()
                {
                    new GarmentSubconDeliveryLetterOut(SubconDeliveryLetterOutGuid, null,null,new Guid(),"","",DateTimeOffset.Now,1,"","",1,"", false)
                });

            Guid SubconDeliveryLetterOutItemGuid = Guid.NewGuid();
            //GarmentSubconDeliveryLetterOutItem garmentSubconDeliveryLetterOutItem = new GarmentSubconDeliveryLetterOutItem(Guid.Empty,new Guid(), 1, new Domain.Shared.ValueObjects.ProductId(1), "code", "name", "remark", "color", 1, new Domain.Shared.ValueObjects.UomId(1), "unit", "fabType");
            //_mockGarmentSubconDeliveryLetterOutItemRepository
            //    .Setup(s => s.Query)
            //    .Returns(new List<GarmentSubconDeliveryLetterOutItemReadModel>()
            //    {
            //        garmentSubconDeliveryLetterOutItem.GetReadModel()
            //    }.AsQueryable());

            // Act
            var result = await unitUnderTest.Get();

            // Assert
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(result));
        }

        [Fact]
        public async Task GetSingle_StateUnderTest_ExpectedBehavior()
        {
            // Arrange
            var unitUnderTest = CreateGarmentSubconDeliveryLetterOutController();

            _mockGarmentSubconDeliveryLetterOutRepository
                .Setup(s => s.Find(It.IsAny<Expression<Func<GarmentSubconDeliveryLetterOutReadModel, bool>>>()))
                .Returns(new List<GarmentSubconDeliveryLetterOut>()
                {
                    new GarmentSubconDeliveryLetterOut(Guid.NewGuid(), null,null,new Guid(),"","",DateTimeOffset.Now,1,"","",1,"", false)
                });

            _mockGarmentSubconDeliveryLetterOutItemRepository
                .Setup(s => s.Find(It.IsAny<Expression<Func<GarmentSubconDeliveryLetterOutItemReadModel, bool>>>()))
                .Returns(new List<GarmentSubconDeliveryLetterOutItem>()
                {
                    new GarmentSubconDeliveryLetterOutItem(Guid.NewGuid(),new Guid(),1,new Domain.Shared.ValueObjects.ProductId(1),"code","name","remark","color",1,new Domain.Shared.ValueObjects.UomId(1),"unit","fabType")
                });


            // Act
            var result = await unitUnderTest.Get(Guid.NewGuid().ToString());

            // Assert
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(result));
        }

        [Fact]
        public async Task Post_StateUnderTest_ExpectedBehavior()
        {
            // Arrange
            var unitUnderTest = CreateGarmentSubconDeliveryLetterOutController();
            PlaceGarmentSubconDeliveryLetterOutCommand command = new PlaceGarmentSubconDeliveryLetterOutCommand();
            command.UENId = 1;
            //_mockGarmentSubconDeliveryLetterOutRepository
            //    .Setup(s => s.Find(It.IsAny<Expression<Func<GarmentSubconDeliveryLetterOutReadModel, bool>>>()))
            //    .Returns(new List<GarmentSubconDeliveryLetterOut>());
            _MockMediator
                .Setup(s => s.Send(It.IsAny<PlaceGarmentSubconDeliveryLetterOutCommand>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(new GarmentSubconDeliveryLetterOut(Guid.NewGuid(), null, null, new Guid(), "", "", DateTimeOffset.Now, 1, "", "", 1, "", false));

            // Act
            var result = await unitUnderTest.Post(command);

            // Assert
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(result));
        }

        [Fact]
        public async Task Post_Throws_Exception()
        {
            // Arrange
            var unitUnderTest = CreateGarmentSubconDeliveryLetterOutController();

            _MockMediator
                .Setup(s => s.Send(It.IsAny<PlaceGarmentSubconDeliveryLetterOutCommand>(), It.IsAny<CancellationToken>()))
                .Throws(new Exception());

            // Act and Assert
            await Assert.ThrowsAsync<Exception>(() => unitUnderTest.Post(It.IsAny<PlaceGarmentSubconDeliveryLetterOutCommand>()));
        }
    }
}
