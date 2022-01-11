﻿using Infrastructure.Domain.Commands;
using System;
using System.Collections.Generic;
using System.Text;

namespace Manufactures.Domain.GarmentSample.SampleCuttingOuts.Commands
{
    public class UpdateDatesGarmentSampleCuttingOutCommand : ICommand<int>
    {
        public UpdateDatesGarmentSampleCuttingOutCommand(List<string> ids, DateTimeOffset date)
        {
            Identities = ids;
            Date = date;
        }

        public List<string> Identities { get; private set; }
        public DateTimeOffset Date { get; private set; }
    }

}
