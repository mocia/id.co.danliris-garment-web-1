﻿using Barebone.Tests;
using FluentAssertions;
using Manufactures.Application.GarmentSewingIns.CommandHandlers;
using Manufactures.Domain.GarmentLoadings;
using Manufactures.Domain.GarmentLoadings.ReadModels;
using Manufactures.Domain.GarmentLoadings.Repositories;
using Manufactures.Domain.GarmentSewingIns;
using Manufactures.Domain.GarmentSewingIns.Commands;
using Manufactures.Domain.GarmentSewingIns.ReadModels;
using Manufactures.Domain.GarmentSewingIns.Repositories;
using Manufactures.Domain.GarmentSewingIns.ValueObjects;
using Manufactures.Domain.GarmentSewingOuts;
using Manufactures.Domain.GarmentSewingOuts.ReadModels;
using Manufactures.Domain.GarmentSewingOuts.Repositories;
using Manufactures.Domain.Shared.ValueObjects;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;


namespace Manufactures.Tests.CommandHandlers.GarmentSewingIns
{
    public class RemoveGarmentSewingInCommandHandlerTests : BaseCommandUnitTest
    {
        private readonly Mock<IGarmentSewingInRepository> _mockSewingInRepository;
        private readonly Mock<IGarmentSewingInItemRepository> _mockSewingInItemRepository;
        private readonly Mock<IGarmentLoadingItemRepository> _mockLoadingItemRepository;
        private readonly Mock<IGarmentSewingOutItemRepository> _mockSewingOutItemRepository;

        public RemoveGarmentSewingInCommandHandlerTests()
        {
            _mockSewingInRepository = CreateMock<IGarmentSewingInRepository>();
            _mockSewingInItemRepository = CreateMock<IGarmentSewingInItemRepository>();
            _mockLoadingItemRepository = CreateMock<IGarmentLoadingItemRepository>();
            _mockSewingOutItemRepository = CreateMock<IGarmentSewingOutItemRepository>();

            _MockStorage.SetupStorage(_mockSewingInRepository);
            _MockStorage.SetupStorage(_mockSewingInItemRepository);
            _MockStorage.SetupStorage(_mockLoadingItemRepository);
            _MockStorage.SetupStorage(_mockSewingOutItemRepository);
        }

        private RemoveGarmentSewingInCommandHandler CreateRemoveGarmentSewingInCommandHandler()
        {
            return new RemoveGarmentSewingInCommandHandler(_MockStorage.Object);
        }

        [Fact]
        public async Task Handle_StateUnderTest_ExpectedBehavior()
        {
            // Arrange
            Guid SewingInGuid = Guid.NewGuid();
            Guid preparingItemGuid = Guid.NewGuid();
            RemoveGarmentSewingInCommandHandler unitUnderTest = CreateRemoveGarmentSewingInCommandHandler();
            CancellationToken cancellationToken = CancellationToken.None;
            RemoveGarmentSewingInCommand RemoveGarmentSewingInCommand = new RemoveGarmentSewingInCommand(SewingInGuid);
            Guid loadingItemGuid = Guid.NewGuid();
            Guid sewingOutItemGuid = Guid.NewGuid();

            _mockSewingInRepository
                .Setup(s => s.Query)
                .Returns(new List<GarmentSewingInReadModel>()
                {
                    new GarmentSewingInReadModel(SewingInGuid)
                }.AsQueryable());
            _mockSewingInItemRepository
                .Setup(s => s.Find(It.IsAny<Expression<Func<GarmentSewingInItemReadModel, bool>>>()))
                .Returns(new List<GarmentSewingInItem>()
                {
                    new GarmentSewingInItem(Guid.Empty, Guid.Empty,sewingOutItemGuid,Guid.Empty,loadingItemGuid, new ProductId(1), null, null, null, new SizeId(1), null, 0, new UomId(1), null, null, 0,1,1)
                });

            //_mockLoadingItemRepository
            //    .Setup(s => s.Query)
            //    .Returns(new List<GarmentLoadingItemReadModel>
            //    {
            //        new GarmentLoadingItemReadModel(loadingItemGuid)
            //    }.AsQueryable());

            _mockSewingOutItemRepository
                .Setup(s => s.Query)
                .Returns(new List<GarmentSewingOutItemReadModel>
                {
                    new GarmentSewingOutItemReadModel(sewingOutItemGuid)
                }.AsQueryable());

            _mockSewingInRepository
                .Setup(s => s.Update(It.IsAny<GarmentSewingIn>()))
                .Returns(Task.FromResult(It.IsAny<GarmentSewingIn>()));
            _mockSewingInItemRepository
                .Setup(s => s.Update(It.IsAny<GarmentSewingInItem>()))
                .Returns(Task.FromResult(It.IsAny<GarmentSewingInItem>()));
            //_mockLoadingItemRepository
            //    .Setup(s => s.Update(It.IsAny<GarmentLoadingItem>()))
            //    .Returns(Task.FromResult(It.IsAny<GarmentLoadingItem>()));
            _mockSewingOutItemRepository
                .Setup(s => s.Update(It.IsAny<GarmentSewingOutItem>()))
                .Returns(Task.FromResult(It.IsAny<GarmentSewingOutItem>()));

            _MockStorage
                .Setup(x => x.Save())
                .Verifiable();

            // Act
            var result = await unitUnderTest.Handle(RemoveGarmentSewingInCommand, cancellationToken);

            // Assert
            result.Should().NotBeNull();
        }
    }
}